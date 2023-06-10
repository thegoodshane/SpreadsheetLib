using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetLib
{
    public static class Code // Implicit number format codes with Ids.
    {
        public const string Number0 = "0";
        public const string Number2 = "0.00";
        public const string Thou0 = "#,##0";
        public const string Thou2 = "#,##0.00";
        public const string Percent0 = "0%";
        public const string Percent2 = "0.00%";
        public const string Scientific2 = "0.00E+00";
        public const string Fraction1 = "# ?/?";
        public const string Fraction2 = "# ??/??";
        public const string MonDayYear = "m/d/yyyy";
        public const string DayMonYear = "d-mmm-yy";
        public const string DayMon = "d-mmm";
        public const string MonYear = "mmm-yy";
        public const string HourMin12 = "h:mm AM/PM";
        public const string HourMinSec12 = "h:mm:ss AM/PM";
        public const string HourMin24 = "h:mm";
        public const string HourMinSec24 = "h:mm:ss";
        public const string DateTime = "m/d/yyyy h:mm";
        public const string Thou0Paren = "#,##0_);(#,##0)";
        public const string Thou0RedParen = "#,##0_);[Red](#,##0)";
        public const string Thou2Paren = "#,##0.00_);(#,##0.00)";
        public const string Thou2RedParen = "#,##0.00_);[Red](#,##0.00)";
        public const string MinSec0 = "mm:ss";
        public const string ElapsedTime = "[h]:mm:ss";
        public const string MinSec1 = "mm:ss.0";
        public const string Scientific1 = "##0.0E+0";
        public const string Text = "@";
    }
}
