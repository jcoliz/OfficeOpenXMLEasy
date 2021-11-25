using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace jcoliz.OfficeOpenXml.Serializer
{
    public interface ISharedStringMap
    {
        string FindSharedStringItem(string id);
    }

    /// <summary>
    /// Reader to deserialize objects from spreadsheets, using Office OpenXML
    /// </summary>
    /// <see href="https://github.com/OfficeDev/Open-XML-SDK"/>
    /// <remarks>
    /// Originally, I used EPPlus. However, that library has terms for commercial use.
    /// </remarks>
    public class SpreadsheetReader : ISpreadsheetReader, ISharedStringMap
    {
        #region ISpreadsheetReader (Public Interface)

        /// <summary>
        /// The names of all the individual sheets
        /// </summary>
        public IEnumerable<string> SheetNames { get; private set; }

        /// <summary>
        /// Open the reader for reading from <paramref name="stream"/>
        /// </summary>
        /// <param name="stream">Where to read from</param>
        public void Open(Stream stream)
        {
            spreadSheet = SpreadsheetDocument.Open(stream, isEditable: false);
            var workbookpart = spreadSheet.WorkbookPart;
            SheetNames = workbookpart.Workbook.Descendants<Sheet>().Select(x => x.Name.Value).ToList();
        }
        
        /// <summary>
        /// Read the sheet named <paramref name="sheetname"/> into items
        /// </summary>
        /// <remarks>
        /// This can be called multiple times on the same open reader
        /// </remarks>
        /// <typeparam name="T">Type of the items to return</typeparam>
        /// <param name="sheetname">Name of sheet. Will be inferred from name of <typeparamref name="T"/> if not supplied.
        /// Will use first sheet in workbook if it's not found.</param>
        /// <param name="exceptproperties">Properties to exclude from the import</param>
        /// <returns>Enumerable of <typeparamref name="T"/> items, OR null if  <paramref name="sheetname"/> is not found</returns>
        public IEnumerable<T> Deserialize<T>(string sheetname = null, IEnumerable<string> exceptproperties = null) where T : class, new()
        {
            // Fill in default name if not specified
            var name = string.IsNullOrEmpty(sheetname) ? typeof(T).Name : sheetname;

            // Find the worksheet

            var workbookpart = spreadSheet.WorkbookPart;
            var matching = workbookpart.Workbook.Descendants<Sheet>().Where(x => x.Name == name);

            if (matching.Any())
            {
                if (matching.Skip(1).Any())
                    throw new ApplicationException($"Ambiguous sheet name. Shreadsheet has multiple sheets matching {name}.");
            }
            else
            {
                matching = workbookpart.Workbook.Descendants<Sheet>();

                if (!matching.Any())
                    return null;
            }

            var sheet = matching.Single();
            WorksheetPart worksheetPart = (WorksheetPart)(workbookpart.GetPartById(sheet.Id));

            // Transform cells into a repository we can work with more easily
            var cells = new CellRepository(worksheetPart.Worksheet.Descendants<Cell>(),this);
            
            // Extract the headers
            var headers = cells.Rows().First().Columns().ToDictionary(c => c.Column, c => c.Value);

            // For each data row
            var result = cells.Rows().Skip(1).Select
            (
                // Create a resulting item
                r => CreateFromDictionary<T>
                (
                    // From a mapping of header to column value
                    r.Columns()
                        .Where
                        (
                            c => !string.IsNullOrEmpty(headers.GetValueOrDefault(c.Column))
                            && exceptproperties?.Any(p => p == headers[c.Column]) != true 
                        )
                        .ToDictionary(c=>headers[c.Column], c=>c.Value)
                )
            );

            return result;
        }

        #endregion

        #region Internals
        /// <summary>
        /// Look up a string from the shared string table part
        /// </summary>
        /// <param name="id">ID for the string, 0-based integer in string form</param>
        /// <exception cref="ApplicationException">
        /// Throws if there is no string table, or if the string can't be found.
        /// </exception>
        /// <returns>The string found</returns>
        public string FindSharedStringItem(string id)
        {
            var shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().SingleOrDefault();

            if (null == shareStringPart)
                throw new ApplicationException("Shared string cell found, but no shared string table!");

            var table = shareStringPart.SharedStringTable;
            var found = table.Skip(Convert.ToInt32(id));
            var result = found.FirstOrDefault()?.InnerText;

            if (null == result)
                throw new ApplicationException($"Unable to find shared string reference for id {id}!");

            return result;
        }

        #endregion

        #region Static Internals

        /// <summary>
        /// Create an object from a <paramref name="dictionary"/> of property values
        /// </summary>
        /// <typeparam name="T">What type of object to create</typeparam>
        /// <param name="dictionary">Dictionary of strings to values, where each key is a property name</param>
        /// <returns>The created object of type <typeparamref name="T"/></returns>
        private static T CreateFromDictionary<T>(Dictionary<string, string> dictionary) where T : class, new()
        {
            var item = new T();

            foreach (var kvp in dictionary)
            {
                var property = typeof(T).GetProperties().Where(x => x.Name == kvp.Key).SingleOrDefault();
                if (null != property?.SetMethod)
                {
                    // Note this excludes typeof(int), which excludes importing
                    // the ID. So if you re-import items you already have, the IDs
                    // will be stripped and ignored. Generally I think that's
                    // good behaviour, but in other implementations, I have done
                    // it the other way where duplicating the ID is a method of
                    // bulk editing.
                    //
                    // ... And this runs up against Pbi #870 :) I now want to
                    // export TransactionID for splits. This means I also need to
                    // export ID for Transactions, so the TransactionID has meaning!

                    if (property.PropertyType == typeof(DateTime))
                    {
                        // By the time datetimes get here, we expect them to be OADates.
                        // If the original source is an actual date type, that should
                        // be adjusted before now.

                        if ( double.TryParse(kvp.Value, out double dvalue) )
                        {
                            var value = DateTime.FromOADate(dvalue);
                            property.SetValue(item, value);
                        }
                    }
                    else if (property.PropertyType == typeof(int))
                    {
                        if (int.TryParse(kvp.Value, out int value))
                            property.SetValue(item, value);
                    }
                    else if (property.PropertyType == typeof(decimal))
                    {
                        if (decimal.TryParse(kvp.Value, out decimal value))
                            property.SetValue(item, value);
                    }
                    else if (property.PropertyType == typeof(bool))
                    {
                        // Bool is represented as 0/1.
                        // But maybe somettimes it will come in as true/false
                        // So I'll deal with each

                        if (int.TryParse(kvp.Value, out int intvalue))
                            property.SetValue(item, intvalue != 0);
                        else
                        {
                            if (bool.TryParse(kvp.Value, out bool value))
                                property.SetValue(item, value);
                        }
                    }
                    else if (property.PropertyType == typeof(string))
                    {
                        var value = kvp.Value?.Trim();
                        if (!string.IsNullOrEmpty(value))
                            property.SetValue(item, value);
                    }
                    else if (property.PropertyType.BaseType == typeof(Enum))
                    {
                        if (Enum.TryParse(property.PropertyType, kvp.Value, out object value))
                            property.SetValue(item, value);
                    }
                }
            }

            return item;
        }

        #endregion

        #region Fields
        SpreadsheetDocument spreadSheet;
        #endregion

        #region IDispose
        private bool disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~NewSpreadsheetReader()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
