using SpreadsheetLib.SpreadsheetML;
using System.Collections.Generic;
using System.Linq;
using System;

namespace SpreadsheetLib
{
    internal class Worksheet : IWorksheet
    {
        private IDictionary<string, Cell> cells = new Dictionary<string, Cell>(); // [cell reference]

        internal Worksheet(string name, int sheetId)
        {
            Name = name;
            SheetId = sheetId;
        }

        internal Worksheet(
            CTWorksheet worksheet,
            CTSheet sheet,
            IList<Style> styles,
            IList<string> sharedStrings)
        {
            // Data validations
            var dataValidations = GetDVs(worksheet.DataValidations);

            // Rows
            if (worksheet.Rows != null) // Rows: null check
            {
                var rowIndex = 1; // Rows: best guess

                foreach (var row in worksheet.Rows) // Rows: loop
                {
                    if (row.RowIndex != null) // Rows: correction
                    {
                        rowIndex = (int)row.RowIndex;
                    }

                    if (row.Cells != null) // Columns: null check
                    {
                        var colIndex = 1; // Columns: best guess

                        foreach (var cell in row.Cells) // Columns: loop
                        {
                            if (cell.Reference != null) // Columns: correction
                            {
                                var address = new Address(cell.Reference);

                                rowIndex = address.RowNumber;

                                colIndex = address.ColumnNumber;
                            }
                            else
                            {
                                cell.Reference = new Address(rowIndex, colIndex).ToString();
                            }

                            cells[cell.Reference] = new Cell(cell, styles, sharedStrings, dataValidations);

                            colIndex++; // Columns: increment
                        }
                    }

                    rowIndex++; // Rows: increment
                }
            }

            // Sheet
            Name = sheet.SheetName;

            SheetId = (int)sheet.SheetTabId;
        }

        public int FirstCol =>
            cells.Keys.Select(_ => new Address(_)).Min(_ => (int?)_.ColumnNumber) ?? 0;

        public int FirstRow =>
            cells.Keys.Select(_ => new Address(_)).Min(_ => (int?)_.RowNumber) ?? 0;

        public bool IsEmpty => FirstRow == 0;

        public int LastCol =>
                    cells.Keys.Select(_ => new Address(_)).Max(_ => (int?)_.ColumnNumber) ?? 0;

        public int LastRow =>
            cells.Keys.Select(_ => new Address(_)).Max(_ => (int?)_.RowNumber) ?? 0;

        public string Name { get; }

        internal string RelationshipId => $"worksheet{SheetId}";

        internal int SheetId { get; }

        public ICell Cell(int rowNumber, int columnNumber)
        {
            var address = new Address(rowNumber, columnNumber);

            var key = address.ToString();

            if (!cells.ContainsKey(key))
            {
                cells[key] = new Cell(address);
            }

            return cells[key];
        }

        public ICell Cell(string address)
        {
            var a = new Address(address);

            return Cell(a.RowNumber, a.ColumnNumber);
        }

        internal static Dictionary<string, DataValidation> GetDVs(IEnumerable<CTDataValidation> ctdvs)
        {
            var dvs = new Dictionary<string, DataValidation>(); // [cell reference]

            if (ctdvs != null)
            {
                foreach (var ctdv in ctdvs)
                {
                    var cellReferences = ctdv.SequenceOfReferences.Split(new[] { ' ' });

                    foreach (var cellReference in cellReferences)
                    {
                        // Handle a range.
                        if (cellReference.Contains(":"))
                        {
                            foreach (var rangeRef in Range.GetCellRefs(cellReference))
                            {
                                dvs[rangeRef] = new DataValidation(ctdv);
                            }
                        }

                        dvs[cellReference] = new DataValidation(ctdv);
                    }
                }
            }

            return dvs;
        }

        internal CTWorksheet Save(IList<Style> styles, IList<string> sharedStrings)
        {
            var rows = new Dictionary<int, CTRow>();

            var dataValidations = new Dictionary<DataValidation, IList<string>>();

            foreach (var cell in cells.Values)
            {
                var rowNumber = cell.Address.RowNumber;

                if (!rows.ContainsKey(rowNumber))
                {
                    rows[rowNumber] = new CTRow()
                    {
                        Cells = new List<CTCell>(),
                        RowIndex = (uint)rowNumber
                    };
                }

                rows[rowNumber].Cells.Add(cell.Save(styles, sharedStrings));

                // Data validations are done here because cells don't reference them.

                // If the data validation has value,
                if (cell.DataValidation.Value != null)
                {
                    var dataValidation = (DataValidation)cell.DataValidation;

                    // If it's not in the dictionary, add it.
                    if (!dataValidations.ContainsKey(dataValidation))
                    {
                        dataValidations[dataValidation] = new List<string>();
                    }

                    // Add this address to the list of cell references.
                    dataValidations[dataValidation].Add(cell.Address.ToString());
                }
            }

            return new CTWorksheet(RelationshipId)
            {
                DataValidations = dataValidations.Select(_ => _.Key.Save(_.Value)).ToList(),
                Rows = rows.Values
            };
        }
    }
}