// <xsd:complexType name="CT_Worksheet">
//   <xsd:sequence>
//     <xsd:element name="sheetData" type="CT_SheetData" minOccurs="1" maxOccurs="1"/>
//     <xsd:element name="dataValidations" type="CT_DataValidations" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
// </xsd:complexType>
// <xsd:complexType name="CT_SheetData">
//   <xsd:sequence>
//     <xsd:element name="row" type="CT_Row" minOccurs="0" maxOccurs="unbounded"/>
//   </xsd:sequence>
// </xsd:complexType>
// <xsd:complexType name="CT_DataValidations">
//   <xsd:sequence>
//     <xsd:element name="dataValidation" type="CT_DataValidation" minOccurs="1" maxOccurs="unbounded"/>
//   </xsd:sequence>
// </xsd:complexType>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>worksheet</summary>
    internal class CTWorksheet
    {
        public CTWorksheet(string relationshipId)
        {
            RelationshipId = relationshipId;
        }

        public CTWorksheet(XElement worksheet, string relationshipId)
        {
            Rows = worksheet.Descendants2("row")
                .Select(_ => new CTRow(_))
                .ToList();

            DataValidations = worksheet.Descendants2("dataValidation")
                .Select(_ => new CTDataValidation(_))
                .ToList();

            RelationshipId = relationshipId;
        }

        public ICollection<CTDataValidation> DataValidations { get; set; }

        public string RelationshipId { get; set; } // Not in schema.

        public ICollection<CTRow> Rows { get; set; }

        public XElement ToXElement(XNamespace ns)
        {
            if (RelationshipId == null)
            {
                throw new InvalidOperationException("A relationship ID is required.");
            }

            var sheetData = new XElement(ns + "sheetData");

            if (Rows != null)
            {
                sheetData.Add(Rows.Select(_ => _.ToXElement(ns)));
            }

            var worksheet = new XElement(ns + "worksheet", sheetData);

            if (DataValidations != null && DataValidations.Count > 0) // minOccurs="1" 
            {
                var dataValidations = DataValidations.Select(_ => _.ToXElement(ns));

                worksheet.Add(new XElement(ns + "dataValidations", dataValidations));
            }

            return worksheet;
        }
    }
}