using SpreadsheetLib.SpreadsheetML;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SpreadsheetLib
{
    public class Workbook : IWorkbook
    {
        public Workbook()
        {
        }

        public Workbook(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open))
            {
                OpenPackage(new OpcPackage(stream));
            }
        }

        public Workbook(Stream stream) => OpenPackage(new OpcPackage(stream));

        public IWorksheets Worksheets { get; } = new Worksheets();

        public void Save(string path)
        {
            using (var stream = new FileStream(path, FileMode.Create))
            {
                Save(stream);
            }
        }

        public void Save(Stream stream) => MakePackage().Save(stream);

        internal static CTStylesheet MakeStylesheet(IList<Style> styles)
        {
            var fills = new List<Fill>()
            {
                new Fill(), // Default,
                new Fill()
                {
                    PatternType = STPatternType.gray125
                }
            };

            var numFmtCodes = new List<string>();

            // Create the cell formats before the fills and numbering formats.
            var cellFormats = styles
                .Select(_ => _.Save(fills, numFmtCodes))
                .ToList();

            var patternFills = fills
                .Select(_ => _.Save())
                .ToList();

            var numberFormats = numFmtCodes
                .Select(_ => new CTNumberFormat((uint)numFmtCodes.IndexOf(_) + 166, _))
                .ToList();

            return new CTStylesheet()
            {
                CellFormats = cellFormats,
                NumberFormats = numberFormats,
                PatternFills = patternFills
            };
        }

        internal OpcPackage MakePackage()
        {
            var package = new OpcPackage()
            {
                Sheets = new List<CTSheet>(),
                Worksheets = new List<CTWorksheet>()
            };

            var styles = new List<Style>()
            {
                new Style() // Default
            };

            var sharedStrings = new List<string>();

            foreach (Worksheet ws in Worksheets)
            {
                package.Sheets.Add(new CTSheet(ws.Name, (uint)ws.SheetId, ws.RelationshipId));

                package.Worksheets.Add(ws.Save(styles, sharedStrings));
            }

            package.Stylesheet = MakeStylesheet(styles);

            package.SharedStringTable = new CTSharedStringTable()
            {
                SharedStrings = sharedStrings
            };

            return package;
        }

        internal void OpenPackage(OpcPackage package)
        {
            var stylesheet = package.Stylesheet;

            var styles = new List<Style>();

            if (stylesheet != null)
            {
                var fills = new List<Fill>();

                if (stylesheet.PatternFills != null)
                {
                    fills = stylesheet.PatternFills
                        .Select(_ => new Fill(_))
                        .ToList();
                }

                var numberFormats = new Dictionary<int, string>();

                if (stylesheet.NumberFormats != null)
                {
                    numberFormats = stylesheet.NumberFormats.ToDictionary(
                        _ => (int)_.NumberFormatId,
                        _ => _.NumberFormatCode);
                }

                if (stylesheet.CellFormats != null)
                {
                    styles = stylesheet.CellFormats
                        .Select(_ => new Style(_, fills, numberFormats))
                        .ToList();
                }
            }

            var sharedStringTable = package.SharedStringTable;

            var sharedStrings = new List<string>();

            if (sharedStringTable != null)
            {
                sharedStrings = sharedStringTable.SharedStrings.ToList();
            }

            foreach (var sheet in package.Sheets.OrderBy(_ => _.SheetTabId))
            {
                var worksheet = package.Worksheets
                    .Single(_ => _.RelationshipId == sheet.RelationshipId);

                var worksheets = (Worksheets)Worksheets;

                worksheets.Add(new Worksheet(worksheet, sheet, styles, sharedStrings));
            }
        }
    }
}