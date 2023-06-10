// <xsd:complexType name="CT_NumFmt">
//   <xsd:attribute name="numFmtId" type="xsd:unsignedInt" use="required"/>
//   <xsd:attribute name="formatCode" type="xsd:string" use="required"/>
// </xsd:complexType>

using System;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"numFmt"</summary>
    internal class CTNumberFormat
    {
        public CTNumberFormat(uint numberFormatId, string numberFormatCode)
        {
            NumberFormatId = numberFormatId;

            NumberFormatCode = numberFormatCode;
        }

        public CTNumberFormat(XElement numberFormat)
        {
            NumberFormatId = uint.Parse(numberFormat.Attribute2("numFmtId").Value);

            NumberFormatCode = numberFormat.Attribute2("formatCode").Value;
        }

        /// <summary>"formatCode", required</summary>
        public string NumberFormatCode { get; set; }

        /// <summary>"numFmtId", required</summary>
        public uint NumberFormatId { get; set; }

        public XElement ToXElement(XNamespace ns)
        {
            var numberFormat = new XElement(ns + "numFmt");

            numberFormat.Add(new XAttribute("numFmtId", NumberFormatId));

            if (NumberFormatCode == null)
            {
                throw new InvalidOperationException("The 'formatCode' attribute is required.");
            }

            numberFormat.Add(new XAttribute("formatCode", NumberFormatCode));

            return numberFormat;
        }
    }
}