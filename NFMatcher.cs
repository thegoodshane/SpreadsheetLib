// Excel en-US format code ABNF (corrected):
// LITERAL-CHAR = REVERSE-SOLIDUS UTF16-ANY 
// LITERAL-STRING = (QUOTATION-MARK 1*UTF16-ANY-WITHOUT-QUOTE QUOTATION-MARK) / 1*LITERAL-CHAR
// LITERAL-CHAR-REPEAT = ASTERISK UTF16-ANY
// LITERAL-CHAR-SPACE = LOW-LINE UTF16-ANY
// LiteralSymbol = "$" / "-" / "+" / "(" / ")" / ":" / " "
// Literal = LiteralSymbol / LITERAL-STRING / LITERAL-CHAR-REPEAT / LITERAL-CHAR-SPACE
// NFPartLocaleID = LEFT-SQUARE-BRACKET DOLLAR-SIGN *ALPHA [HYPHEN-MINUS 3*8DIGIT-HEXADECIMAL] RIGHT-SQUARE-BRACKET
// NFPartYear = 2(SMALL-LETTER-Y) / 4(SMALL-LETTER-Y)
// NFPartMonth = 1*5(SMALL-LETTER-M)
// NFPartDay = 1*4(SMALL-LETTER-D)
// NFPartHour = 1*2(SMALL-LETTER-H)
// NFPartMinute = 1*2(SMALL-LETTER-M)
// NFPartSecond = 1*2(SMALL-LETTER-S) 
// NFPartAbsHour = LEFT-SQUARE-BRACKET 1*SMALL-LETTER-H RIGHT-SQUARE-BRACKET
// NFPartAbsMinute = LEFT-SQUARE-BRACKET 1*SMALL-LETTER-M RIGHT-SQUARE-BRACKET
// NFPartAbsSecond = LEFT-SQUARE-BRACKET 1*SMALL-LETTER-S RIGHT-SQUARE-BRACKET
// NFPartSubSecond = INTL-CHAR-DECIMAL-SEP 1*3DIGIT-ZERO
// NFAbsTimeToken = NFPartAbsHour / NFPartAbsMinute / NFPartAbsSecond
// NFDateTimeToken = NFPartYear / NFPartMonth / NFPartDay / NFPartHour / NFPartMinute / NFPartSecond / NFAbsTimeToken
// NFPartSubSecond = INTL-CHAR-DECIMAL-SEP 1*3DIGIT-ZERO
// AMPM = (CAPITAL-LETTER-A CAPITAL-LETTER-M SOLIDUS CAPITAL-LETTER-P CAPITAL-LETTER-M) / "A/P"
// NFDateTime = *NFPartLocaleID (1*(NFDateTimeToken) *(NFDateTimeToken / NFPartSubSecond / CHAR-DATE-SEP / AMPM / Literal))

using System.Text.RegularExpressions;

namespace SpreadsheetLib
{
    internal static class NFMatcher
    {
        private const string locale =   // NFPartLocaleID
            @"\[\$\w*(\-[0-9A-Fa-f]{3,8})?\]";

        private const string token =    // NFDateTimeToken
            @"y{2}|y{4}" +              // NFPartYear              
            @"|m{1,5}" +                // or NFPartMonth
            @"|d{1,4}" +                // or NFPartDay
            @"|h{1,2}" +                // or NFPartHour
            @"|m{1,2}" +                // or NFPartMinute
            @"|s{1,2}" +                // or NFPartSecond
            @"|\[h+\]" +                // or NFPartAbsHour
            @"|\[m+\]" +                // or NFPartAbsMinute
            @"|\[s+\]";                 // or NFPartAbsSecond

        private const string literal =  // Literal
            @"[\$\+\(\) -:]" +          // LiteralSymbol
            @"|""[^""]+""|(\\.)+" +     // or LITERAL-STRING
            @"|\*." +                   // or LITERAL-CHAR-REPEAT
            @"|_.";                     // or LITERAL-CHAR-SPACE

        private const string datetime = // NFDateTime
            @"^(" + locale + ")*" +     // *NFPartLocaleID
            @"((" + token + ")+" +      // 1*(NFDateTimeToken)
            @"(" + token +              // NFDateTimeToken
            @"|\.(0){1,3}" +            // or NFPartSubSecond
            @"|/" +                     // or CHAR-DATE-SEP 
            @"|AM/PM|A/P" +             // or AMPM         
            @"|" + literal +            // or Literal
            @")*)$";

        internal static bool IsDateTime(string numberFormatCode) =>
            new Regex(datetime).IsMatch(numberFormatCode);
    }
}
