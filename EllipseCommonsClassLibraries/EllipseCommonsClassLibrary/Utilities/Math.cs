using System;
using System.Collections.Generic;
using System.Linq;

namespace CommonsClassLibrary.Utilities
{
    public class Math
    {
        /// <summary>
        ///     Indica si un número está dentro del rango aceptado con respecto a otro. Si compareNumber está dentro de los valores
        ///     de baseNumber +/- (baseNumber * threshold)
        /// </summary>
        /// <param name="baseNumber">object: Número base de operación </param>
        /// <param name="compareNumber">object: Número a comparar</param>
        /// <param name="threshold">object: Rango de comparación (Ej.: 0.1, 0.3, 1, 1.3, 2)</param>
        /// <returns>
        ///     bool: true si compareNumber está dentro del rango threshold de baseNumber (Ej. baseNumber: 6, compareNumber:
        ///     4. Si threshold es 0.5 es true, si threshold es 0.3 es false
        /// </returns>
        public static bool InThreshold(object baseNumber, object compareNumber, object threshold)
        {
            if (baseNumber == null || compareNumber == null)
                throw new NullReferenceException(
                    "No se puede realizar la operación. Ingreso debe ser diferente de nulo");

            var num1 = Convert.ToDecimal(baseNumber);
            var num2 = Convert.ToDecimal(compareNumber);
            var th = Convert.ToDecimal(threshold);
            var abs = System.Math.Abs(num1 - num2);

            return num1 * th >= abs;
        }

        public static string ToOrdinal(long number)
        {
            if (number < 0) return number.ToString();
            var rem = number % 100;
            if (rem >= 11 && rem <= 13) return number + "th";

            switch (number % 10)
            {
                case 1:
                    return number + "st";
                case 2:
                    return number + "nd";
                case 3:
                    return number + "rd";
                default:
                    return number + "th";
            }
        }

        public static string ToOrdinal(int number)
        {
            return ToOrdinal((long)number);
        }

        public static string ToOrdinal(string number)
        {
            if (string.IsNullOrEmpty(number)) return number;

            var dict = new Dictionary<string, string>
            {
                {"zero", "zeroth"},
                {"nought", "noughth"},
                {"one", "first"},
                {"two", "second"},
                {"three", "third"},
                {"four", "fourth"},
                {"five", "fifth"},
                {"six", "sixth"},
                {"seven", "seventh"},
                {"eight", "eighth"},
                {"nine", "ninth"},
                {"ten", "tenth"},
                {"eleven", "eleventh"},
                {"twelve", "twelfth"},
                {"thirteen", "thirteenth"},
                {"fourteen", "fourteenth"},
                {"fifteen", "fifteenth"},
                {"sixteen", "sixteenth"},
                {"seventeen", "seventeenth"},
                {"eighteen", "eighteenth"},
                {"nineteen", "nineteenth"},
                {"twenty", "twentieth"},
                {"thirty", "thirtieth"},
                {"forty", "fortieth"},
                {"fifty", "fiftieth"},
                {"sixty", "sixtieth"},
                {"seventy", "seventieth"},
                {"eighty", "eightieth"},
                {"ninety", "ninetieth"},
                {"hundred", "hundredth"},
                {"thousand", "thousandth"},
                {"million", "millionth"},
                {"billion", "billionth"},
                {"trillion", "trillionth"},
                {"quadrillion", "quadrillionth"},
                {"quintillion", "quintillionth"}
            };


            // rough check whether it's a valid number
            var temp = number.ToLower().Trim().Replace(" and ", " ");
            var words = temp.Split(new[] {' ', '-'}, StringSplitOptions.RemoveEmptyEntries);

            if (words.Any(word => !dict.ContainsKey(word)))
                return number;

            // extract last word
            number = number.TrimEnd().TrimEnd('-');
            var index = number.LastIndexOfAny(new[] {' ', '-'});
            var last = number.Substring(index + 1);

            // make replacement and maintain original capitalization
            if (last == last.ToLower())
                last = dict[last];
            else if (last == last.ToUpper())
                last = dict[last.ToLower()].ToUpper();
            else
                last = last.ToLower();
            last = char.ToUpper(dict[last][0]) + dict[last].Substring(1);

            return number.Substring(0, index + 1) + last;
        }
    }
}