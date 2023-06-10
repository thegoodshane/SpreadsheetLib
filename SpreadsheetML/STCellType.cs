// <xsd:simpleType name="ST_CellType">
//   <xsd:restriction base="xsd:string">
//     <xsd:enumeration value="b"/>
//     <xsd:enumeration value="d"/>
//     <xsd:enumeration value="n"/>
//     <xsd:enumeration value="e"/>
//     <xsd:enumeration value="s"/>
//     <xsd:enumeration value="str"/>
//     <xsd:enumeration value="inlineStr"/>
//   </xsd:restriction>
// </xsd:simpleType>

namespace SpreadsheetLib.SpreadsheetML
{
    internal enum STCellType
    {
        /// <summary>Number</summary>
        n,
        /// <summary>Boolean</summary>
        b,
        /// <summary>ISO 8601 Date</summary>
        d,
        /// <summary>Error</summary>
        e,
        /// <summary>Shared String</summary>
        s,
        /// <summary>Formula String</summary>
        str,
        /// <summary>Inline String</summary>
        inlineStr
    }
}