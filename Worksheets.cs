using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System;

namespace SpreadsheetLib
{
    internal class Worksheets : IWorksheets
    {
        private IList<Worksheet> worksheets = new List<Worksheet>();

        public IWorksheet Add(string name)
        {
            if (worksheets.Any(_ => _.Name == name))
            {
                throw new ArgumentException("Worksheet name must be unique.");
            }

            var sheetId = worksheets.Count + 1;

            var worksheet = new Worksheet(name, sheetId);

            Add(worksheet);

            return worksheet;
        }

        internal void Add(Worksheet worksheet)
        {
            worksheets.Add(worksheet);
        }

        public IEnumerator<IWorksheet> GetEnumerator() => worksheets.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}