// Interfaces inspired by ClosedXML.

using System;
using System.Collections.Generic;
using System.IO;

namespace SpreadsheetLib
{
    public interface IWorkbook
    {
        IWorksheets Worksheets { get; } // Default: new

        void Save(string path);

        void Save(Stream stream);
    }

    public interface IWorksheets : IEnumerable<IWorksheet>
    {
        IWorksheet Add(string name);
    }

    public interface IWorksheet
    {
        bool IsEmpty { get; }

        /// <summary>The first column used. Returns 0 if no cells exist.</summary>
        int FirstCol { get; } // Default: N/A

        /// <summary>The first row used. Returns 0 if no cells exist.</summary>
        int FirstRow { get; } // Default: N/A

        /// <summary>The last column used. Returns 0 if no cells exist.</summary>
        int LastCol { get; } // Default: N/A

        /// <summary>The last row used. Returns 0 if no cells exist.</summary>
        int LastRow { get; } // Default: N/A

        string Name { get; }

        ICell Cell(int rowNumber, int columnNumber);

        ICell Cell(string address);
    }

    public interface ICell
    {
        /// <summary>
        /// Supports: string, double, bool, DateTime
        /// </summary>
        object Value { get; set; } // Default: null

        IStyle Style { get; } // Default: new

        IDataValidation DataValidation { get; } // Default: new
    }

    public interface IAddress
    {
        int ColumnNumber { get; } // Default: N/A

        int RowNumber { get; } // Default: N/A
    }

    public interface IDataValidation
    {
        string Value { get; } // Default: null

        void List(string list);
    }

    public interface IStyle
    {
        IFill Fill { get; } // Default: new

        INumberFormat NumberFormat { get; } // Default: new
    }

    public interface IFill
    {
        Color? Color { get; set; } // Default: null
    }

    public interface INumberFormat
    {
        string Code { get; set; } // Default: null
    }
}
