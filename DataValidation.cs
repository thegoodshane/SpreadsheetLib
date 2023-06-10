using SpreadsheetLib.SpreadsheetML;
using System;
using System.Collections.Generic;

namespace SpreadsheetLib
{
    // Equatable because data validations are shared by cells.
    internal class DataValidation : IDataValidation, IEquatable<DataValidation>
    {
        internal DataValidation()
        {
        }

        internal DataValidation(CTDataValidation dataValidation)
        {
            if (dataValidation.DataValidationType == STDataValidationType.list)
            {
                Value = dataValidation.Formula1?.Trim(new[] { '"' });
            }
        }

        public string Value { get; private set; }

        #region Equatable

        public bool Equals(DataValidation other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Value, other.Value);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((DataValidation) obj);
        }

        public override int GetHashCode()
        {
            return (Value != null ? Value.GetHashCode() : 0);
        }

        public static bool operator ==(DataValidation left, DataValidation right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(DataValidation left, DataValidation right)
        {
            return !Equals(left, right);
        }
        
        #endregion Equatable

        public void List(string list)
        {
            Value = list;
        }

        internal CTDataValidation Save(IEnumerable<string> cellReferences)
        {
            var sequenceOfReferences = String.Join(" ", cellReferences);

            return new CTDataValidation(sequenceOfReferences)
            {
                AllowBlank = true,
                Formula1 = $"\"{Value}\"",
                DataValidationType = STDataValidationType.list
            };
        }
    }
}