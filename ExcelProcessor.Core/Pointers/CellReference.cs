using ExcelProcessor.Abstractions.Pointers;
using System.Globalization;

namespace ExcelProcessor.Core.Pointers
{
    public class CellReference : ICellReference
    {
        public int Row { get; private set; }
        public string Column { get; private set; }

        public CellReference(int row, string column)
        {
            if (string.IsNullOrEmpty(column))
                throw new ArgumentNullException(column);

            string columnValue = column.ToUpper();
            foreach (char character in columnValue)
            {
                if (character < 'A' || character > 'Z')
                    throw new ArgumentException("Invalid column name", nameof(column));
            }

            Row = row;
            Column = columnValue;
        }

        public ICellReference NextColumn()
        {
            return new CellReference(Row, NextColumn(Column));
        }

        public ICellReference NextRow()
        {
            return new CellReference(Row + 1, Column);
        }

        private static string NextColumn(string column)
        {
            IEnumerable<char> characters = column.ToCharArray().Reverse();
            string nextColumn = string.Empty;
            bool sumCharacter = true;
            foreach (char character in characters)
            {
                if (sumCharacter)
                {
                    sumCharacter = false;
                    char nextCharacter = (char)(character + 1);
                    if (nextCharacter > 'Z')
                    {
                        nextCharacter = 'A';
                        sumCharacter = true;
                    }
                    nextColumn = nextCharacter + nextColumn;
                }
                else
                    nextColumn = character + nextColumn;
            }
            if (sumCharacter)
                nextColumn = "A" + nextColumn;
            return nextColumn;
        }

        public string ToExcelString()
        {
            return $"{Column}{Row}";
        }

        public override string ToString()
        {
            return ToExcelString();
        }

        public int GetColumnIndex()
        {
            int columnNumber = -1;
            int mulitplier = 1;

            foreach (char c in Column.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * (c - 64);
                mulitplier = mulitplier * 26;
            }
            return columnNumber + 1;
        }

        public string ToDoubleChange(decimal number)
        {
            NumberFormatInfo nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            return number.ToString(nfi);
        }
    }
}
