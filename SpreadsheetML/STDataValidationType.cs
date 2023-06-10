// <xsd:simpleType name="ST_DataValidationType">
//   <xsd:restriction base="xsd:string">
//     <xsd:enumeration value="none"/>
//     <xsd:enumeration value="whole"/>
//     <xsd:enumeration value="decimal"/>
//     <xsd:enumeration value="list"/>
//     <xsd:enumeration value="date"/>
//     <xsd:enumeration value="time"/>
//     <xsd:enumeration value="textLength"/>
//     <xsd:enumeration value="custom"/>
//   </xsd:restriction>
// </xsd:simpleType>

namespace SpreadsheetLib.SpreadsheetML
{
    internal enum STDataValidationType
    {
        none,
        whole,
        @decimal,
        list,
        date,
        time,
        textLength,
        custom
    }
}