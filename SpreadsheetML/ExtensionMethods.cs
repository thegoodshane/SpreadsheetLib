using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.IO;

namespace SpreadsheetLib.SpreadsheetML
{
    internal static class ExtensionMethods
    {
        public static XAttribute Attribute2<T>(this T source, string localName)
            where T : XElement
        {
            return source.Attributes().SingleOrDefault(_ => _.Name.LocalName == localName);
        }

        public static IEnumerable<XElement> Descendants2<T>(this T source, string localName)
            where T : XContainer
        {
            return source.Descendants().Where(e => e.Name.LocalName == localName);
        }

        public static IEnumerable<XElement> Elements2<T>(this T source, string localName) // 1145659
                            where T : XContainer
        {
            return source.Elements().Where(e => e.Name.LocalName == localName);
        }

        public static string GetFileName(this string path)
        {
            return Path.GetFileName(path);
        }

        public static uint? ToNullableUInt(this string s)
        {
            uint u;

            return uint.TryParse(s, out u) ? u : (uint?)null;
        }

        public static bool ToBool(this string xsdBoolean)
        {
            return xsdBoolean == "1" || xsdBoolean == "true";
        }

        public static int AddOrIndex<T>(this IList<T> list, T item)
        {
            if (!list.Contains(item))
            {
                list.Add(item);
            }

            return list.IndexOf(item);
        }
    }
}