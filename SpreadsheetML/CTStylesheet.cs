// <xsd:complexType name="CT_Stylesheet">
//   <xsd:sequence>
//     <xsd:element name="numFmts" type="CT_NumFmts" minOccurs="0" maxOccurs="1"/>
//     <xsd:element name="fills" type="CT_Fills" minOccurs="0" maxOccurs="1"/>
//     <xsd:element name="cellXfs" type="CT_CellXfs" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
// </xsd:complexType>
// <xsd:complexType name="CT_NumFmts">
//   <xsd:sequence>
//     <xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="unbounded"/>
//   </xsd:sequence>
// </xsd:complexType>
// <xsd:complexType name="CT_Fills">
//   <xsd:sequence>
//     <xsd:element name="fill" type="CT_Fill" minOccurs="0" maxOccurs="unbounded"/>
//   </xsd:sequence>
// </xsd:complexType>
// <xsd:complexType name="CT_Fill">
//   <xsd:choice minOccurs="1" maxOccurs="1">
//     <xsd:element name="patternFill" type="CT_PatternFill" minOccurs="0" maxOccurs="1"/>
//   </xsd:choice>
// </xsd:complexType>
// <xsd:complexType name="CT_CellXfs">
//   <xsd:sequence>
//     <xsd:element name="xf" type="CT_Xf" minOccurs="1" maxOccurs="unbounded"/>
//   </xsd:sequence>
// </xsd:complexType>

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>styleSheet</summary>
    internal class CTStylesheet
    {
        public CTStylesheet()
        {
        }

        public CTStylesheet(XElement stylesheet)
        {
            // "numFmt" appears in "numFmts" and elsewhere.
            var numberFormats = stylesheet.Elements2("numFmts").SingleOrDefault();

            if (numberFormats != null)
            {
                NumberFormats = numberFormats.Elements2("numFmt")
                    .Select(_ => new CTNumberFormat(_))
                    .ToList();
            }

            PatternFills = stylesheet.Descendants2("patternFill")
                .Select(_ => new CTPatternFill(_))
                .ToList();

            // "xf" elements appear in both the "cellStyleXfs" and "cellXfs" sections.
            var cellFormats = stylesheet.Elements2("cellXfs").SingleOrDefault();

            if (cellFormats != null)
            {
                CellFormats = cellFormats.Elements2("xf")
                    .Select(_ => new CTCellFormat(_))
                    .ToList();
            }
        }

        public ICollection<CTCellFormat> CellFormats { get; set; }

        public ICollection<CTNumberFormat> NumberFormats { get; set; }

        public ICollection<CTPatternFill> PatternFills { get; set; }

        // A font, fill, and border are required if implicitly or explicitly referenced by a cell format.
        // Excel requires the order: number formats, fonts, fills, borders, cell formats.
        public XElement ToXElement(XNamespace ns)
        {
            var stylesheet = new XElement(ns + "styleSheet");

            if (CellFormats?.Count > 0)
            {
                if (NumberFormats?.Count > 0)
                {
                    stylesheet.Add(new XElement(ns + "numFmts", NumberFormats.Select(_ => _.ToXElement(ns))));
                }

                stylesheet.Add(new XElement(ns + "fonts", new XElement(ns + "font")));

                var fills = new XElement(ns + "fills");

                if (PatternFills?.Count > 0)
                {
                    fills.Add(PatternFills.Select(_ => new XElement(ns + "fill", _.ToXElement(ns))));
                }
                else
                {
                    fills.Add(new XElement(ns + "fill"));
                }

                stylesheet.Add(fills);

                stylesheet.Add(new XElement(ns + "borders", new XElement(ns + "border")));

                stylesheet.Add(new XElement(ns + "cellXfs", CellFormats.Select(_ => _.ToXElement(ns))));
            }

            return stylesheet;
        }
    }
}