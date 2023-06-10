using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    internal class OpcPackage
    {
        private const string
            appCT = "application/xml",
            relCT = "application/vnd.openxmlformats-package.relationships+xml",
            wbCT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            wsCT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            ssCT = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
            sstCT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

        private static readonly XNamespace
            pkgRelNS = "http://schemas.openxmlformats.org/package/2006/relationships",
            refRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            ssRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            sstRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            typeNS = "http://schemas.openxmlformats.org/package/2006/content-types",
            wbRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            wsRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            xlNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        public OpcPackage()
        {
        }

        // Part names are used instead of full names because some
        // apps use relative paths and others use absolute.
        public OpcPackage(Stream stream)
        {
            using (var archive = new ZipArchive(stream))
            {
                var parts = archive.Entries;

                // Entry point -> package relationships part.
                var packageRelsPart = parts.Single(_ => _.FullName == "_rels/.rels");

                // Package relationships part -> package relationships.
                var packageRels = XElement.Load(packageRelsPart.Open());

                // Package relationships -> workbook name.
                var workbookName = packageRels.Elements2("Relationship")
                    .Single(_ => _.Attribute2("Type").Value.EndsWith("officeDocument"))
                    .Attribute2("Target")
                    .Value
                    .GetFileName();

                // Workbook name -> workbook part.
                var workbookPart = parts.Single(_ => _.Name == workbookName);

                // Workbook part -> workbook.
                var workbook = XElement.Load(workbookPart.Open());

                // Workbook -> sheets.
                Sheets = workbook.Descendants2("sheet")
                    .Select(_ => new CTSheet(_))
                    .ToList();

                // Workbook part -> workbook folder.
                //var workbookFolder = workbookPart.FullName.Replace(workbookPart.Name, "");

                // Workbook part -> workbook relationships name.
                var workbookRelsName = $"{workbookPart.Name}.rels";

                // Workbook relationships name -> workbook relationships part.
                var workbookRelsPart = parts.Single(_ => _.Name == workbookRelsName);

                // Workbook relationships part -> workbook relationships.
                var workbookRels = XElement.Load(workbookRelsPart.Open());

                // Workbook relationships -> worksheet names / IDs.
                var wsNamesIds = workbookRels.Elements2("Relationship")
                    .Where(_ => _.Attribute2("Type").Value.EndsWith("worksheet"))
                    .ToDictionary(
                        _ => _.Attribute2("Target").Value.GetFileName(),
                        _ => _.Attribute2("Id").Value);

                // Worksheet names / IDs -> worksheets.
                Worksheets = parts.Where(_ => wsNamesIds.Keys.Contains(_.Name))
                    .Select(_ => new CTWorksheet(XElement.Load(_.Open()), wsNamesIds[_.Name]))
                    .ToList();

                // Workbook relationships -> shared strings name.
                var sharedStringsName = workbookRels.Elements2("Relationship")
                    .SingleOrDefault(_ => _.Attribute2("Type").Value.EndsWith("sharedStrings"))
                    ?.Attribute2("Target")
                    ?.Value
                    ?.GetFileName();

                // Shared strings name -> shared strings part.
                var sharedStringsPart = parts.SingleOrDefault(_ => _.Name == sharedStringsName);

                // Shared strings part -> shared string table.
                if (sharedStringsPart != null)
                {
                    SharedStringTable =
                        new CTSharedStringTable(XElement.Load(sharedStringsPart.Open()));
                }

                // Workbook relationships -> styles name.
                var stylesName = workbookRels.Elements2("Relationship")
                    .SingleOrDefault(_ => _.Attribute2("Type").Value.EndsWith("styles"))
                    ?.Attribute2("Target")
                    ?.Value
                    ?.GetFileName();

                // Styles name -> styles part.
                var stylesPart = parts.SingleOrDefault(_ => _.Name == stylesName);

                // Styles part -> stylesheet.
                if (stylesPart != null)
                {
                    Stylesheet = new CTStylesheet(XElement.Load(stylesPart.Open()));
                }
            }
        }

        public CTSharedStringTable SharedStringTable { get; set; }

        public ICollection<CTSheet> Sheets { get; set; }

        public CTStylesheet Stylesheet { get; set; }

        public ICollection<CTWorksheet> Worksheets { get; set; }

        public void Save(Stream stream)
        {
            if (Sheets.Count == 0)
            {
                throw new InvalidOperationException("At least one sheet is required.");
            }

            if (Sheets.Count != Worksheets.Count)
            {
                throw new InvalidOperationException("The number of sheets and worksheets must match.");
            }

            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                // Package relationships
                var packageRelsPart = archive.CreateEntry("_rels/.rels");

                using (var writer = new StreamWriter(packageRelsPart.Open()))
                {
                    MakePackageRels().Save(writer);
                }

                // Workbook
                var workbookPart = archive.CreateEntry("xl/workbook.xml");

                using (var writer = new StreamWriter(workbookPart.Open()))
                {
                    MakeWorkbook().Save(writer);
                }

                // Content types
                var contentTypesPart = archive.CreateEntry("[Content_Types].xml");

                using (var writer = new StreamWriter(contentTypesPart.Open()))
                {
                    MakeContentTypes().Save(writer);
                }

                // Workbook relationships
                var workbookRelsPart = archive.CreateEntry("xl/_rels/workbook.xml.rels");

                using (var writer = new StreamWriter(workbookRelsPart.Open()))
                {
                    MakeWorkbookRels().Save(writer);
                }

                // Worksheets
                foreach (var sheet in Sheets)
                {
                    var worksheetPart = archive.CreateEntry($"xl/worksheets/sheet{sheet.SheetTabId}.xml");

                    using (var writer = new StreamWriter(worksheetPart.Open()))
                    {
                        Worksheets.Single(_ => _.RelationshipId == sheet.RelationshipId)
                            .ToXElement(xlNS)
                            .Save(writer);
                    }
                }

                // Styles
                if (Stylesheet != null)
                {
                    var stylesPart = archive.CreateEntry("xl/styles.xml");

                    using (var writer = new StreamWriter(stylesPart.Open()))
                    {
                        Stylesheet.ToXElement(xlNS).Save(writer);
                    }
                }

                // Shared strings
                if (SharedStringTable != null)
                {
                    var sharedStringsPart = archive.CreateEntry("xl/sharedStrings.xml");

                    using (var writer = new StreamWriter(sharedStringsPart.Open()))
                    {
                        SharedStringTable.ToXElement(xlNS).Save(writer);
                    }
                }
            }
        }

        internal XElement MakePackageRels()
        {
            return
                new XElement(pkgRelNS + "Relationships",
                    new XElement(pkgRelNS + "Relationship",
                        new XAttribute("Id", "workbook"),
                        new XAttribute("Target", "xl/workbook.xml"),
                        new XAttribute("Type", wbRelNS)));
        }

        internal XElement MakeWorkbook()
        {
            return
                new XElement(xlNS + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", refRelNS),
                    new XElement(xlNS + "sheets", Sheets.Select(_ => _.ToXElement(xlNS, refRelNS))));
        }

        internal XElement MakeContentTypes()
        {
            var types =
                new XElement(typeNS + "Types",
                    new XElement(typeNS + "Default",
                        new XAttribute("ContentType", appCT),
                        new XAttribute("Extension", "xml")),
                    new XElement(typeNS + "Default",
                        new XAttribute("ContentType", relCT),
                        new XAttribute("Extension", "rels")),
                    new XElement(typeNS + "Override",
                        new XAttribute("ContentType", wbCT),
                        new XAttribute("PartName", "/xl/workbook.xml")));

            foreach (var sheet in Sheets)
            {
                types.Add(
                    new XElement(typeNS + "Override",
                        new XAttribute("ContentType", wsCT),
                        new XAttribute("PartName", $"/xl/worksheets/sheet{sheet.SheetTabId}.xml"))
                    );
            }

            if (Stylesheet != null)
            {
                types.Add(
                    new XElement(typeNS + "Override",
                        new XAttribute("ContentType", ssCT),
                        new XAttribute("PartName", "/xl/styles.xml"))
                    );
            }

            if (SharedStringTable != null)
            {
                types.Add(
                    new XElement(typeNS + "Override",
                        new XAttribute("ContentType", sstCT),
                        new XAttribute("PartName", "/xl/sharedStrings.xml"))
                    );
            }

            return types;
        }

        internal XElement MakeWorkbookRels()
        {
            var workbookRels = new XElement(pkgRelNS + "Relationships");

            foreach (var sheet in Sheets)
            {
                workbookRels.Add(
                    new XElement(pkgRelNS + "Relationship",
                        new XAttribute("Id", sheet.RelationshipId),
                        new XAttribute("Target", $"worksheets/sheet{sheet.SheetTabId}.xml"),
                        new XAttribute("Type", wsRelNS))
                    );
            }

            if (Stylesheet != null)
            {
                workbookRels.Add(
                    new XElement(pkgRelNS + "Relationship",
                        new XAttribute("Id", "styles"),
                        new XAttribute("Target", "styles.xml"),
                        new XAttribute("Type", ssRelNS))
                    );
            }

            if (SharedStringTable != null)
            {
                workbookRels.Add(
                    new XElement(pkgRelNS + "Relationship",
                        new XAttribute("Id", "sharedStrings"),
                        new XAttribute("Target", "sharedStrings.xml"),
                        new XAttribute("Type", sstRelNS))
                    );
            }

            return workbookRels;
        }
    }
}