using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace jcoliz.OfficeOpenXml.Serializer
{
    /// <summary>
    /// Holds a table of OpenXml cells in a format we can easily iterate over
    /// </summary>
    /// <remarks>
    /// Note that this works as a fully rectangular 2D table. If one row has
    /// 25 columns, then ALL rows have 25 columns, even if the last 24 are
    /// nulls.
    /// </remarks>
    internal class CellRepository
    {
        private readonly Dictionary<string, Cell> _dictionary;
        private readonly int _maxrow;
        private readonly ISharedStringMap _stringMap;

        public uint MaxCols { get; private set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="cells">Which cells to work with</param>
        /// <param name="stringMap">Where to get strings</param>
        public CellRepository(IEnumerable<Cell> cells, ISharedStringMap stringMap)
        {
            _dictionary = cells.ToDictionary(x => x.CellReference.Value, x => x);
            _stringMap = stringMap;

            // Determine extent of cells

            // Note that rows are 1-based, and columns are 0-based, to make them easier to convert to/from letters
            var regex = new Regex(@"([A-Za-z]+)(\d+)");
            var matches = _dictionary.Keys.Select(x => regex.Match(x).Groups);

            // Rows are the second matching group in the cell identifier
            _maxrow = matches.Max(x => Convert.ToInt32(x[2].Value));

            // Columns are the first matching group
            MaxCols = matches.Max(x => ColNumberFor(x[1].Value));
        }

        /// <summary>
        /// All the rows in this table of cells
        /// </summary>
        /// <returns></returns>
        public IEnumerable<RepositoryRow> Rows() =>
            Enumerable.Range(1, _maxrow).Select(r => new RepositoryRow(this, r));

        /// <summary>
        /// Find a single cell value in the repository
        /// </summary>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public RepositoryValue this[uint col, int row]
        {
            get
            {
                RepositoryValue result = null; 
                var cell = _dictionary.GetValueOrDefault(ColNameFor(col) + row);
                if (null != cell)
                {
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        result = new RepositoryValue() { Column = col, Value = _stringMap.FindSharedStringItem(cell.CellValue?.Text) };

                    else if (!string.IsNullOrEmpty(cell.CellValue?.Text))
                        result = new RepositoryValue() { Column = col, Value = cell.CellValue.Text };
                }
                return result;
            }
        }

        /// <summary>
        /// Convert string column name to integer index
        /// </summary>
        /// <param name="colname">Base 26-style column name, e.g. "AF"</param>
        /// <returns>0-based integer column number, e.g. "A" = 0</returns>
        private static uint ColNumberFor(IEnumerable<char> colname)
        {
            if (colname == null || !colname.Any())
                return 0;

            var last = (uint)colname.Last() - (uint)'A';
            var others = ColNumberFor(colname.SkipLast(1));

            return last + 26U * (1 + others);
        }

        /// <summary>
        /// Convert column number to spreadsheet name
        /// </summary>
        /// <param name="colnumber">0-based integer column number, e.g. "A" = 0</param>
        /// <returns>Base 26-style column name, e.g. "AF"</returns>
        private static string ColNameFor(uint number)
        {
            if (number < 26)
                return new string(new char[] { (char)((int)'A' + number) });
            else
                return ColNameFor((number / 26) - 1) + ColNameFor(number % 26);
        }
    }

    /// <summary>
    /// A single row in a CellRepository
    /// </summary>
    internal class RepositoryRow
    {
        /// <summary>
        /// Which repository contains this row
        /// </summary>
        private readonly CellRepository _repository;

        /// <summary>
        /// What row is this, starting with 1
        /// </summary>
        private readonly int _row;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="repository">Which repository holds this row</param>
        /// <param name="row">Which row# are we, starting from 1</param>
        public RepositoryRow(CellRepository repository, int row)
        {
            _repository = repository;
            _row = row;
        }

        /// <summary>
        /// Obtain all the cell values within the columns
        /// </summary>
        /// <returns></returns>
        public IEnumerable<RepositoryValue> Columns() =>
            Enumerable.Range(0, (int)_repository.MaxCols + 1).Select(x => _repository[(uint)x, _row]);

        /// <summary>
        /// The value of a certain column
        /// </summary>
        /// <param name="col">Which column, starting with 0</param>
        /// <returns></returns>
        public RepositoryValue this[uint col] => _repository[col,_row];
    }

    /// <summary>
    /// A single cell value
    /// </summary>
    internal class RepositoryValue
    {
        /// <summary>
        /// Which column was the value found in, starting from 0
        /// </summary>
        public uint Column { get; set; }
        
        /// <summary>
        /// What value was found in the column
        /// </summary>
        public string Value { get; set; }
    }
}
