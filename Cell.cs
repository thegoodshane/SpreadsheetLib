using SpreadsheetLib.SpreadsheetML;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace SpreadsheetLib
{
    internal class Cell : ICell
    {
        internal Cell(Address address)
        {
            Address = address;
        }

        internal Cell(
            CTCell cell,
            IList<Style> styles,
            IList<string> sharedStrings,
            IDictionary<string, DataValidation> dataValidations)
        {
            // cell.Reference was already null checked.
            Address = new Address(cell.Reference);

            // Set style before value.
            Style = (Style)styles[(int)cell.StyleIndex].Clone();

            Value = cell.CellValue;

            switch (cell.CellDataType)
            {
                case STCellType.n: // "double-precision floating point number"
                    if (double.TryParse(cell.CellValue, out var d))
                    {
                        Value = d;

                        if (((NumberFormat)Style.NumberFormat).IsDateTime)
                        {
                            Value = DateTime.FromOADate(d);
                        }
                    }
                    break;
                case STCellType.b: // "1 to specify true and 0 to specify false"
                    Value = cell.CellValue.ToBool();
                    break;
                case STCellType.d:
                    if (DateTime.TryParse(cell.CellValue, null, DateTimeStyles.RoundtripKind, out var dt))
                    {
                        Value = dt;
                    }
                    break;
                case STCellType.s: // "zero-based index into the shared string table"
                    if (int.TryParse(cell.CellValue, out var i))
                    {
                        Value = sharedStrings[i];
                    }
                    break;
                case STCellType.inlineStr:
                    Value = cell.InlineString;
                    break;
                default: // Error (e) or formula result (str)
                    break;
            }

            DataValidation dataValidation;

            if (dataValidations.TryGetValue(Address.ToString(), out dataValidation))
            {
                DataValidation = dataValidation;
            }
        }

        public IDataValidation DataValidation { get; } = new DataValidation();

        public IStyle Style { get; } = new Style();

        public object Value { get; set; }

        internal IAddress Address { get; }

        internal CTCell Save(IList<Style> styles, IList<string> sharedStrings)
        {
            var cell = new CTCell()
            {
                Reference = Address.ToString()
            };

            string defaultFormatCode = null;

            switch (Value)
            {
                case string s:
                    cell.CellValue = sharedStrings.AddOrIndex(s).ToString();
                    cell.CellDataType = STCellType.s;
                    break;
                case bool b:
                    cell.CellValue = (b ? 1 : 0).ToString();
                    cell.CellDataType = STCellType.b;
                    break;
                case DateTime dt:
                    cell.CellValue = dt.ToOADate().ToString();
                    defaultFormatCode = Code.DateTime;
                    break;
                case double d:
                    cell.CellValue = d.ToString();
                    defaultFormatCode = Code.Number2;
                    break;
                case int i:
                    cell.CellValue = i.ToString();
                    defaultFormatCode = Code.Number0;
                    break;
                default:
                    if (Value != null)
                    {
                        cell.CellValue = Value.ToString();
                        cell.CellDataType = STCellType.s;
                    }
                    break;
            }

            // If the format hasn't been set but a default is necessary,
            if (Style.NumberFormat.Code == null && defaultFormatCode != null)
            {
                Style.NumberFormat.Code = defaultFormatCode;
            }

            // Styles are added last because the number format code may have changed.
            cell.StyleIndex = (uint)styles.AddOrIndex((Style)Style);

            return cell;
        }
    }
}