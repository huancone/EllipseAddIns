using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharedClassLibrary.Utilities
{
    public class IxConstantInteger
    {
        public int Value { get; set; }

        public IxConstantInteger()
        {
            Value = default(int);
        }
        public IxConstantInteger(int value)
        {
            Value = value;
        }

        public bool Equals(IxConstantInteger iConversionConstant)
        {
            return Value == iConversionConstant.Value;
        }

        public static implicit operator IxConstantInteger(int value)
        {
            return new IxConstantInteger { Value = value };
        }

        public static implicit operator int(IxConstantInteger value)
        {
            return value.Value;
        }
    }

    public class IxConversionConstant : IxConstantInteger
    {
        public static readonly IxConversionConstant DefaultNormal = 0;
        public static readonly IxConversionConstant DefaultEmptyOnly = 1;
        public static readonly IxConversionConstant DefaultNullAndEmpty = 2;
        public static readonly IxConversionConstant DefaultAllNonNullValues = 3;

        public static bool IsValidDefaultNullableConstant(int defaultConstant)
        {
            return IsValidDefaultNullableConstant(new IxConversionConstant(defaultConstant));
        }
        public static bool IsValidDefaultNullableConstant(IxConversionConstant defaultConstant)
        {
            if (defaultConstant == null)
                return true;
            if (defaultConstant.Equals(DefaultNormal))
                return true;
            if (defaultConstant.Equals(DefaultEmptyOnly))
                return true;
            if (defaultConstant.Equals(DefaultNullAndEmpty))
                return true;
            if (defaultConstant.Equals(DefaultAllNonNullValues))
                return true;

            return false;
        }

        public static implicit operator IxConversionConstant(int value)
        {
            return new IxConversionConstant { Value = value };
        }
        public IxConversionConstant()
        {

        }
        public IxConversionConstant(int value)
        {
            Value = value;
        }
    }
}
