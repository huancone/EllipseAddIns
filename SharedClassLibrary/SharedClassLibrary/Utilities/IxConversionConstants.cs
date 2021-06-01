using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedClassLibrary.Utilities
{

    public enum IxConversionConstant
    {
        DefaultNormal = 0,
        DefaultEmptyOnly = 1,
        DefaultNullAndEmpty = 2,
        DefaultAllNonNullValues = 3
    }

    public static class IxConversionConstantMethods
    {
        public static bool IsValidDefaultNullableConstant(this IxConversionConstant cc)
        {
            return Enum.IsDefined((typeof(IxConversionConstant)), cc);
        }

    }
}


