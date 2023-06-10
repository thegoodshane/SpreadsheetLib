// <xsd:complexType name="CT_Row">
//   <xsd:sequence>
//     <xsd:element name="c" type="CT_Cell" minOccurs="0" maxOccurs="unbounded"/>
//   </xsd:sequence>
//   <xsd:attribute name="r" type="xsd:unsignedInt" use="optional"/>
// </xsd:complexType>

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"row"</summary>
    internal class CTRow
    {
        public CTRow()
        {
        }

        public CTRow(XElement row)
        {
            Cells = row.Elements2("c")
                .Select(_ => new CTCell(_))
                .ToList();

            RowIndex = row.Attribute2("r")
                ?.Value
                ?.ToNullableUInt();
        }

        public ICollection<CTCell> Cells { get; set; }

        /// <summary>"r"</summary>
        public uint? RowIndex { get; set; }

        public XElement ToXElement(XNamespace ns)
        {
            var row = new XElement(ns + "row");

            if (Cells != null)
            {
                row.Add(Cells.Select(_ => _.ToXElement(ns)));
            }

            if (RowIndex != null)
            {
                row.Add(new XAttribute("r", RowIndex));
            }

            return row;
        }
    }
}