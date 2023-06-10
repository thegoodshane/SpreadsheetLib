// <xsd:complexType name="CT_Sst">
//   <xsd:sequence>
//     <xsd:element name="si" type="CT_Rst" minOccurs="0" maxOccurs="unbounded"/>
//   </xsd:sequence>
// </xsd:complexType>
// <xsd:complexType name="CT_Rst">
//   <xsd:sequence>
//     <xsd:element name="t" type="xsd:string" minOccurs="0" maxOccurs="1"/>
//   </xsd:sequence>
// </xsd:complexType>

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace SpreadsheetLib.SpreadsheetML
{
    /// <summary>"sst"</summary>
    internal class CTSharedStringTable
    {
        public CTSharedStringTable()
        {
        }

        public CTSharedStringTable(XElement sst)
        {
            // Preserve order and index.
            SharedStrings = sst.Elements2("si")
                .Select(_ => _.Elements2("t").SingleOrDefault()?.Value)
                .ToList();
        }

        public ICollection<string> SharedStrings { get; set; }

        public XElement ToXElement(XNamespace ns)
        {
            var sharedStringTable = new XElement(ns + "sst");

            if (SharedStrings != null)
            {
                foreach (var sharedString in SharedStrings)
                {
                    var stringItem = (sharedString == null) ?
                        new XElement(ns + "si") :
                        new XElement(ns + "si", new XElement(ns + "t", sharedString));

                    sharedStringTable.Add(stringItem);
                }
            }

            return sharedStringTable;
        }
    }
}