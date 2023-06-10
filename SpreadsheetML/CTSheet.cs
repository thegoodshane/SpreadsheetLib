// <xsd:complexType name="CT_Sheet">
//   <xsd:attribute name="name" type="xsd:string" use="required"/>
//   <xsd:attribute name="sheetId" type="xsd:unsignedInt" use="required"/>
//   <xsd:attribute name="id" type="xsd:string" use="required"/>
// </xsd:complexType>

using System;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"sheet"</summary>
    internal class CTSheet
    {
        public CTSheet(string sheetName, uint sheetTabId, string relationshipId)
        {
            SheetName = sheetName;

            SheetTabId = sheetTabId;

            RelationshipId = relationshipId;
        }

        public CTSheet(XElement sheet)
        {
            SheetName = sheet.Attribute2("name").Value;

            SheetTabId = uint.Parse(sheet.Attribute2("sheetId").Value);

            RelationshipId = sheet.Attribute2("id").Value;
        }

        /// <summary>"id"</summary>
        public string RelationshipId { get; set; }

        /// <summary>"name", unique, required</summary>
        public string SheetName { get; set; }

        /// <summary>"sheetId", unique, required</summary>
        public uint SheetTabId { get; set; }

        public XElement ToXElement(XNamespace xlNS, XNamespace refRelNS)
        {
            var sheet = new XElement(xlNS + "sheet");

            if (SheetName == null)
            {
                throw new InvalidOperationException("The 'name' attribute is required.");
            }

            sheet.Add(new XAttribute("name", SheetName));

            if (SheetTabId == 0)
            {
                throw new InvalidOperationException("The 'sheetId' attribute is required.");
            }

            sheet.Add(new XAttribute("sheetId", SheetTabId));

            if (RelationshipId == null)
            {
                throw new InvalidOperationException("The 'id' attribute is required.");
            }

            sheet.Add(new XAttribute(refRelNS + "id", RelationshipId));

            return sheet;
        }
    }
}