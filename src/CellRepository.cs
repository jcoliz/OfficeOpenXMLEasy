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
    internal class CellRepository
    {
        private readonly Dictionary<string, Cell> _dictionary;
        private readonly int _maxrow;
        private readonly ISharedStringMap _stringMap;

        public uint MaxCols { get; private set; }

        public CellRepository(IEnumerable<Cell> cells, ISharedStringMap stringMap)
        {
            _dictionary = cells.ToDictionary(x => x.CellReference.Value, x => x);
            _stringMap = stringMap;

            // Determine extent of cells

            // Note that rows are 1-based, and columns are 0-based, to make them easier to convert to/from letters
            var regex = new Regex(@"([A-Za-z]+)(\d+)");
            var matches = _dictionary.Keys.Select(x => regex.Match(x).Groups);
            _maxrow = matches.Max(x => Convert.ToInt32(x[2].Value));
            MaxCols = matches.Max(x => ColNumberFor(x[1].Value));
        }

        public IEnumerable<CellRepositoryRow> Rows()
        {
            return Enumerable.Range(1, _maxrow).Select(r => new CellRepositoryRow(this, r));
        }

        public string this[uint col, int row]
        {
            get
            {
                string result = null;
                var cell = _dictionary.GetValueOrDefault(ColNameFor(col) + row);
                if (null != cell)
                {
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        result = _stringMap.FindSharedStringItem(cell.CellValue?.Text);

                    else if (!string.IsNullOrEmpty(cell.CellValue?.Text))
                        result = cell.CellValue.Text;
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

    internal class CellRepositoryRow
    {
        /// <summary>
        /// Which repository contains this row
        /// </summary>
        private readonly CellRepository _repository;

        /// <summary>
        /// What row is this, starting with 1
        /// </summary>
        private readonly int _row;

        public CellRepositoryRow(CellRepository repository, int row)
        {
            _repository = repository;
            _row = row;
        }

        public IEnumerable<CellRepositoryValue> Columns()
        {
            return Enumerable.Range(0, (int)_repository.MaxCols + 1).Select(x => new CellRepositoryValue() { Column = (uint)x, Value = _repository[(uint)x, _row] });
        }
    }

    internal class CellRepositoryValue
    {
        public uint Column { get; set; }
        public string Value { get; set; }
    }
}
