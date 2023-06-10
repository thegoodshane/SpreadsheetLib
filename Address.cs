using System;

namespace SpreadsheetLib
{
    internal class Address : IAddress
    {
        internal Address(string cellReference)
        {
            int rowIndex = cellReference.IndexOfAny("0123456789".ToCharArray());

            string columnName = cellReference.Substring(0, rowIndex);

            ColumnNumber = ColumnMap.Instance[columnName];

            RowNumber = Int32.Parse(cellReference.Substring(rowIndex));
        }

        internal Address(int rowNumber, int columnNumber)
        {
            if (rowNumber < 1 || columnNumber < 1)
            {
                throw new ArgumentException("Numbers can't be less than 1.");
            }

            ColumnNumber = columnNumber;

            RowNumber = rowNumber;
        }

        public int ColumnNumber { get; }

        public int RowNumber { get; }

        public override string ToString() => ColumnMap.Instance[ColumnNumber] + RowNumber;
    }
}