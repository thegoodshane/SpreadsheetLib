// <xsd:complexType name="CT_DataValidation">
//   <xsd:sequence>
//     <xsd:element name="formula1" type="xsd:string" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
//   <xsd:attribute name="type" type="ST_DataValidationType" use="optional" default="none"/>
//   <xsd:attribute name="allowBlank" type="xsd:boolean" use="optional" default="false"/>
//   <xsd:attribute name="sqref" type="ST_Sqref" use="required"/>
// </xsd:complexType>
// <xsd:simpleType name="ST_Sqref">
//   <xsd:list itemType="xsd:string"/>
// </xsd:simpleType>

using System;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"dataValidation"</summary>
    internal class CTDataValidation
    {
        private STDataValidationType dataValidationType;

        public CTDataValidation(string sequenceOfReferences)
        {
            SequenceOfReferences = sequenceOfReferences;
        }

        public CTDataValidation(XElement dataValidation)
        {
            Formula1 = dataValidation.Elements2("formula1")
                .SingleOrDefault()
                ?.Value;

            Enum.TryParse(dataValidation.Attribute2("type")?.Value, out dataValidationType);

            AllowBlank = dataValidation.Attribute2("allowBlank")
                ?.Value
                ?.ToBool() ?? false;

            SequenceOfReferences = dataValidation.Attribute2("sqref").Value;
        }

        /// <summary>"allowBlank"</summary>
        public bool AllowBlank { get; set; }

        /// <summary>"type"</summary>
        public STDataValidationType DataValidationType
        {
            get { return dataValidationType; }

            set { dataValidationType = value; }
        }

        /// <summary>"formula1"</summary>
        public string Formula1 { get; set; }

        /// <summary>"sqref", required</summary>
        public string SequenceOfReferences { get; set; }

        public XElement ToXElement(XNamespace ns)
        {
            var dataValidation = new XElement(ns + "dataValidation");

            if (Formula1 != null)
            {
                dataValidation.Add(new XElement(ns + "formula1", Formula1));
            }

            if (DataValidationType != STDataValidationType.none)
            {
                dataValidation.Add(new XAttribute("type", DataValidationType));
            }

            if (AllowBlank != false)
            {
                dataValidation.Add(new XAttribute("allowBlank", AllowBlank));
            }

            if (SequenceOfReferences == null)
            {
                throw new InvalidOperationException("The 'sqref' attribute is required.");
            }

            dataValidation.Add(new XAttribute("sqref", SequenceOfReferences));

            return dataValidation;
        }
    }
}