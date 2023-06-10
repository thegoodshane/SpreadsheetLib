// <xsd:complexType name="CT_Cell">
//   <xsd:sequence>
//     <xsd:element name="v" type="xsd:string" minOccurs="0" maxOccurs="1"/>
//     <xsd:element name="is" type="CT_Rst" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
//   <xsd:attribute name="r" type="xsd:string" use="optional"/>
//   <xsd:attribute name="s" type="xsd:unsignedInt" use="optional" default="0"/>
//   <xsd:attribute name="t" type="ST_CellType" use="optional" default="n"/>
// </xsd:complexType>
// <xsd:complexType name="CT_Rst">
//   <xsd:sequence>
//     <xsd:element name="t" type="xsd:string" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
// </xsd:complexType>

using System;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"c"</summary>
    internal class CTCell
    {
        private STCellType cellDataType;

        private uint styleIndex;

        public CTCell()
        {
        }

        public CTCell(XElement cell)
        {
            CellValue = cell.Elements2("v")
                .SingleOrDefault()
                ?.Value;

            InlineString = cell.Elements2("is")
                .SingleOrDefault()
                ?.Elements2("t")
                ?.SingleOrDefault()
                ?.Value;

            Reference = cell.Attribute2("r")?.Value;

            uint.TryParse(cell.Attribute2("s")?.Value, out styleIndex);

            Enum.TryParse(cell.Attribute2("t")?.Value, out cellDataType);
        }

        /// <summary>"t"</summary>
        public STCellType CellDataType
        {
            get { return cellDataType; }

            set { cellDataType = value; }
        }

        /// <summary>"v"</summary>
        public string CellValue { get; set; }

        /// <summary>"is"</summary>
        public string InlineString { get; set; }

        /// <summary>"r"</summary>
        public string Reference { get; set; }

        /// <summary>"s"</summary>
        public uint StyleIndex
        {
            get { return styleIndex; }

            set { styleIndex = value; }
        }

        public XElement ToXElement(XNamespace ns)
        {
            var cell = new XElement(ns + "c");

            if (CellValue != null)
            {
                cell.Add(new XElement(ns + "v", CellValue));
            }

            if (InlineString != null)
            {
                throw new NotImplementedException();
            }

            if (Reference != null)
            {
                cell.Add(new XAttribute("r", Reference));
            }

            if (StyleIndex != 0)
            {
                cell.Add(new XAttribute("s", StyleIndex));
            }

            if (CellDataType != STCellType.n)
            {
                cell.Add(new XAttribute("t", CellDataType));
            }

            return cell;
        }
    }
}