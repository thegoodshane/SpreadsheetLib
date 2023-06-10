// <xsd:simpleType name="ST_PatternType">
//   <xsd:restriction base="xsd:string">
//     <xsd:enumeration value="none"/>
//     <xsd:enumeration value="solid"/>
//     <xsd:enumeration value="mediumGray"/>
//     <xsd:enumeration value="darkGray"/>
//     <xsd:enumeration value="lightGray"/>
//     <xsd:enumeration value="darkHorizontal"/>
//     <xsd:enumeration value="darkVertical"/>
//     <xsd:enumeration value="darkDown"/>
//     <xsd:enumeration value="darkUp"/>
//     <xsd:enumeration value="darkGrid"/>
//     <xsd:enumeration value="darkTrellis"/>
//     <xsd:enumeration value="lightHorizontal"/>
//     <xsd:enumeration value="lightVertical"/>
//     <xsd:enumeration value="lightDown"/>
//     <xsd:enumeration value="lightUp"/>
//     <xsd:enumeration value="lightGrid"/>
//     <xsd:enumeration value="lightTrellis"/>
//     <xsd:enumeration value="gray125"/>
//     <xsd:enumeration value="gray0625"/>
//   </xsd:restriction>
// </xsd:simpleType>

namespace SpreadsheetLib.SpreadsheetML
{
    internal enum STPatternType
    {
        /// <summary>
        /// A pattern of 'none' overrides colors and means the cell has no fill.
        /// </summary>
        none,
        /// <summary>
        /// When solid is specified, the foreground color is the only color rendered.
        /// </summary>
        solid,
        mediumGray,
        darkGray,
        lightGray,
        darkHorizontal,
        darkVertical,
        darkDown,
        darkUp,
        darkGrid,
        darkTrellis,
        lightHorizontal,
        lightVertical,
        lightDown,
        lightUp,
        lightGrid,
        lightTrellis,
        gray125,
        gray0625
    }
}