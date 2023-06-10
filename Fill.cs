using SpreadsheetLib.SpreadsheetML;
using System;
using System.Globalization;

namespace SpreadsheetLib
{
    // Equatable because fills are shared by styles.
    internal class Fill : IFill, IEquatable<Fill>, ICloneable
    {
        internal Fill()
        {
        }

        // Import: ARGB string -> RGB string -> int -> Color
        internal Fill(CTPatternFill patternFill)
        {
            if (patternFill.PatternType == STPatternType.solid)
            {
                string argb = patternFill.ForegroundColor;

                if (argb != null)
                {
                    string rgb = argb.Substring(2);

                    int color = int.Parse(rgb, NumberStyles.HexNumber);

                    Color = (Color)color;
                }
            }
        }

        public Color? Color { get; set; }

        // Remains here only for exporting the gray 125 fill required by Excel.
        internal STPatternType PatternType { get; set; }

        #region Equatable

        public static bool operator !=(Fill left, Fill right)
        {
            return !Equals(left, right);
        }

        public static bool operator ==(Fill left, Fill right)
        {
            return Equals(left, right);
        }

        public bool Equals(Fill other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return Color == other.Color && PatternType == other.PatternType;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Fill)obj);
        }

        public override int GetHashCode()
        {
            int hashCode;

            // Prevent null and black colors from having the same hash code.
            if (Color == null)
            {
                hashCode = -1;
            }
            else
            {
                hashCode = Color.GetHashCode();
            }

            unchecked
            {
                return (hashCode * 397) ^ (int)PatternType;
            }
        }
        #endregion Equatable

        public object Clone() => MemberwiseClone();

        // Export: Color -> int -> RGB string -> ARGB string
        internal CTPatternFill Save()
        {
            var patternFill = new CTPatternFill();

            if (Color != null)
            {
                int color = (int)Color;

                string rgb = color.ToString("X6"),
                    argb = "FF" + rgb;

                patternFill.ForegroundColor = argb;
                patternFill.PatternType = STPatternType.solid;
            }
            else
            {
                patternFill.PatternType = PatternType;
            }

            return patternFill;
        }
    }
}