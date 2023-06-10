using System;

namespace SpreadsheetLib
{
    // Equatable because number formats are shared by styles.
    internal class NumberFormat : INumberFormat, IEquatable<NumberFormat>, ICloneable
    {
        internal NumberFormat()
        {
        }

        internal NumberFormat(string code)
        {
            Code = code;

            var firstSection = Code.Split(new[] { ';' })[0];

            IsDateTime = NFMatcher.IsDateTime(firstSection);
        }

        public string Code { get; set; }

        internal bool IsDateTime { get; }

        public object Clone() => MemberwiseClone();

        #region Equatable

        public static bool operator !=(NumberFormat left, NumberFormat right)
        {
            return !Equals(left, right);
        }

        public static bool operator ==(NumberFormat left, NumberFormat right)
        {
            return Equals(left, right);
        }

        public bool Equals(NumberFormat other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Code, other.Code);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((NumberFormat)obj);
        }

        public override int GetHashCode()
        {
            return (Code != null ? Code.GetHashCode() : 0);
        }

        #endregion Equatable
    }
}