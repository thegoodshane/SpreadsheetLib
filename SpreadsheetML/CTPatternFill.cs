// <xsd:complexType name="CT_PatternFill">
//   <xsd:sequence>
//     <xsd:element name="fgColor" type="CT_Color" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
//   <xsd:attribute name="patternType" type="ST_PatternType" use="optional" default="none"/>
// </xsd:complexType>
// <xsd:complexType name="CT_Color">
//   <xsd:attribute name="rgb" type="xsd:hexBinary" use="optional"/>
// </xsd:complexType>

using System;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"patternFill"</summary>
    internal class CTPatternFill
    {
        private STPatternType patternType;

        public CTPatternFill()
        {
        }

        public CTPatternFill(XElement patternFill)
        {
            ForegroundColor = patternFill.Elements2("fgColor")
                .SingleOrDefault()
                ?.Attribute2("rgb")
                ?.Value;

            Enum.TryParse(patternFill.Attribute2("patternType")?.Value, out patternType);
        }

        /// <summary>"fgColor"</summary>
        public string ForegroundColor { get; set; }

        /// <summary>"patternType"</summary>
        public STPatternType PatternType
        {
            get { return patternType; }

            set { patternType = value; }
        }

        public XElement ToXElement(XNamespace ns)
        {
            var patternFill = new XElement(ns + "patternFill");

            if (ForegroundColor != null)
            {
                patternFill.Add(
                    new XElement(ns + "fgColor", 
                        new XAttribute("rgb", ForegroundColor))
                    );
            }

            if (PatternType != STPatternType.none)
            {
                patternFill.Add(new XAttribute("patternType", PatternType));
            }

            return patternFill;
        }
    }
}