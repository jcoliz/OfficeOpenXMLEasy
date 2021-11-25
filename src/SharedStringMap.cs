using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace jcoliz.OfficeOpenXml.Serializer
{
    /// <summary>
    /// Wrapper for OpenXml shared string table
    /// </summary>
    internal class SharedStringMap : ISharedStringMap
    {
        private readonly SharedStringTable _table;

        public SharedStringMap(SharedStringTablePart part)
        {
            if (null == part)
                throw new ApplicationException("Shared string cell found, but no shared string table!");

            _table = part.SharedStringTable;
        }

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
            var found = _table.Skip(Convert.ToInt32(id));
            var result = found.FirstOrDefault()?.InnerText;

            if (null == result)
                throw new ApplicationException($"Unable to find shared string reference for id {id}!");

            return result;
        }
    }
}
