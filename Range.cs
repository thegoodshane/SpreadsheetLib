using System;
using System.Collections.Generic;

namespace SpreadsheetLib
{
    public static class Range
    {
        public static IEnumerable<string> GetCellRefs(string range)
        {
            var cellRefs = range.Split(new[] { ':' });

            Address from, to;

            try // This is not in the spec.
            {
                from = new Address(cellRefs[0]);
                to = new Address(cellRefs[1]);

                if (from.RowNumber > to.RowNumber || from.ColumnNumber > to.ColumnNumber)
                {
                    throw new Exception();
                }
            }
            catch
            {
                yield break;
            }

            for (int r = from.RowNumber; r <= to.RowNumber; r++)
            {
                for (int c = from.ColumnNumber; c <= to.ColumnNumber; c++)
                {
                    yield return new Address(r, c).ToString();
                }
            }
        }
    }
}
