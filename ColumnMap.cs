using System.Collections.Generic;

namespace SpreadsheetLib
{
    internal class ColumnMap
    {
        // Excel has columns from A (1) to XFD (16384).
        private const int maxColumns = 16384;

        private static ColumnMap instance;

        private IDictionary<string, int> nameNumber = new Dictionary<string, int>(maxColumns);

        private IDictionary<int, string> numberName = new Dictionary<int, string>(maxColumns);

        private ColumnMap()
        {
            var baseChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

            for (int i = 1; i <= maxColumns; i++)
            {
                var name = IntToString(i - 1, baseChars);

                numberName[i] = name;
                nameNumber[name] = i;
            }
        }

        internal static ColumnMap Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ColumnMap();
                }

                return instance;
            }
        }

        internal string this[int index] => numberName[index];

        internal int this[string index] => nameNumber[index];

        private static string IntToString(int value, char[] baseChars) // 923771
        {
            string result = string.Empty;
            int targetBase = baseChars.Length;

            do
            {
                result = baseChars[value % targetBase] + result;
                value = (value / targetBase) - 1;
            }
            while (value > -1);

            return result;
        }
    }
}