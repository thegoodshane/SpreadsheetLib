// <xsd:complexType name="CT_Xf">
//   <xsd:attribute name="numFmtId" type="xsd:unsignedInt" use="optional"/>
//   <xsd:attribute name="fillId" type="xsd:unsignedInt" use="optional"/>
// </xsd:complexType>

using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"xf"</summary>
    internal class CTCellFormat
    {
        public CTCellFormat()
        {
        }

        public CTCellFormat(XElement cellFormat)
        {
            NumberFormatId = cellFormat.Attribute2("numFmtId")
                ?.Value
                ?.ToNullableUInt();

            FillId = cellFormat.Attribute2("fillId")
                ?.Value
                ?.ToNullableUInt();
        }

        /// <summary>"fillId"</summary>
        public uint? FillId { get; set; }

        /// <summary>"numFmtId"</summary>
        public uint? NumberFormatId { get; set; }

        public XElement ToXElement(XNamespace ns)
        {
            var cellFormat = new XElement(ns + "xf");

            if (NumberFormatId != null)
            {
                cellFormat.Add(new XAttribute("numFmtId", NumberFormatId));
            }

            if (FillId != null)
            {
                cellFormat.Add(new XAttribute("fillId", FillId));
            }

            return cellFormat;
        }
    }
}