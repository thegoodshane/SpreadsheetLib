using SpreadsheetLib.SpreadsheetML;
using System;
using System.Collections.Generic;
using static SpreadsheetLib.Code;

namespace SpreadsheetLib
{
    // Cell format alias
    // Equatable because styles are shared by cells.
    internal class Style : IStyle, IEquatable<Style>, ICloneable
    {
        #region ImpliedNumberFormats

        private static readonly IDictionary<string, int> codeId =
            new Dictionary<string, int>
            {
                [Number0] = 1,
                [Number2] = 2,
                [Thou0] = 3,
                [Thou2] = 4,
                [Percent0] = 9,
                [Percent2] = 10,
                [Scientific2] = 11,
                [Fraction1] = 12,
                [Fraction2] = 13,
                [MonDayYear] = 14,
                [DayMonYear] = 15,
                [DayMon] = 16,
                [MonYear] = 17,
                [HourMin12] = 18,
                [HourMinSec12] = 19,
                [HourMin24] = 20,
                [HourMinSec24] = 21,
                [Code.DateTime] = 22,
                [Thou0Paren] = 37,
                [Thou0RedParen] = 38,
                [Thou2Paren] = 39,
                [Thou2RedParen] = 40,
                [MinSec0] = 45,
                [ElapsedTime] = 46,
                [MinSec1] = 47,
                [Scientific1] = 48,
                [Text] = 49,
            };

        private static readonly IDictionary<int, string> idCode =
            new Dictionary<int, string>
            {
                [1] = Number0,
                [2] = Number2,
                [3] = Thou0,
                [4] = Thou2,
                [9] = Percent0,
                [10] = Percent2,
                [11] = Scientific2,
                [12] = Fraction1,
                [13] = Fraction2,
                [14] = MonDayYear,
                [15] = DayMonYear,
                [16] = DayMon,
                [17] = MonYear,
                [18] = HourMin12,
                [19] = HourMinSec12,
                [20] = HourMin24,
                [21] = HourMinSec24,
                [22] = Code.DateTime,
                [37] = Thou0Paren,
                [38] = Thou0RedParen,
                [39] = Thou2Paren,
                [40] = Thou2RedParen,
                [45] = MinSec0,
                [46] = ElapsedTime,
                [47] = MinSec1,
                [48] = Scientific1,
                [49] = Text
            };
        #endregion

        internal Style()
        {
        }

        internal Style(
            CTCellFormat cellFormat,
            IList<Fill> fills,
            IDictionary<int, string> styleNumberFormats)
        {
            if (cellFormat.FillId != null)
            {
                int fillId = (int)cellFormat.FillId;

                Fill = fills[fillId];
            }

            if (cellFormat.NumberFormatId != null)
            {
                int formatId = (int)cellFormat.NumberFormatId;

                string formatCode;

                // If the format ID is not found in the implied number formats,  
                if (!idCode.TryGetValue(formatId, out formatCode))
                {
                    // search in the style number formats.
                    styleNumberFormats.TryGetValue(formatId, out formatCode);
                }

                // If found, replace the default number format.
                if (formatCode != null)
                {
                    NumberFormat = new NumberFormat(formatCode);
                }
            }
        }

        public IFill Fill { get; private set; } = new Fill();

        // Virtual for unit test mocking.
        public virtual INumberFormat NumberFormat { get; private set; } = new NumberFormat();

        #region Equatable

        public static bool operator !=(Style left, Style right)
        {
            return !Equals(left, right);
        }

        public static bool operator ==(Style left, Style right)
        {
            return Equals(left, right);
        }

        public bool Equals(Style other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(Fill, other.Fill) && Equals(NumberFormat, other.NumberFormat);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Style)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Fill != null ? Fill.GetHashCode() : 0) * 397)
                    ^ (NumberFormat != null ? NumberFormat.GetHashCode() : 0);
            }
        }
        #endregion Equatable

        public object Clone()
        {
            ICloneable fill = (ICloneable)Fill,
                numberFormat = (ICloneable)NumberFormat;

            return new Style()
            {
                Fill = (Fill)fill.Clone(),
                NumberFormat = (NumberFormat)numberFormat.Clone()
            };
        }

        internal CTCellFormat Save(IList<Fill> fills, IList<string> numberFormatCodes)
        {
            int numberFormatId = 0;

            if (NumberFormat.Code != null)
            {
                // If the code isn't implied, add it to the list.
                if (!codeId.TryGetValue(NumberFormat.Code, out numberFormatId))
                {
                    numberFormatId = numberFormatCodes.AddOrIndex(NumberFormat.Code) + 166;
                }
            }

            int fillId = fills.AddOrIndex((Fill)Fill);

            //return new CTCellFormat()
            //{
            //    FillId = (uint)fills.AddOrIndex((Fill)Fill),
            //    NumberFormatId = (uint)numberFormatId
            //};

            return new CTCellFormat()
            {
                FillId = (fillId == 0) ? null : (uint?)fillId,
                NumberFormatId = (numberFormatId == 0) ? null : (uint?)numberFormatId
            };
        }
    }
}