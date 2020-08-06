using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EllipseCommonsClassLibrary.Utilities
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
