using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EllipseCommonsClassLibrary.Classes;

namespace EllipseCommonsClassLibrary.Utilities
{
    public class MyUtilities : SharedClassLibrary.Utilities.MyUtilities
    {


        /// <summary>
        ///     Obtiene una lista de tipo string a partir de la llave y valor del listado de keyValuePairList
        /// </summary>
        /// <param name="ellipseCodeItemsList">
        ///     List(EllipseCodeItem{string, string}): Listado tipo EllipseCodeItem con los datos de
        ///     llaves y valores
        /// </param>
        /// <param name="separator">string: separador de llave/descripción (Ej. " - ", "-", "/")</param>
        /// <returns>string: List{code - description}</returns>
        public static List<string> GetCodeList(List<EllipseCodeItem> ellipseCodeItemsList, string separator = " - ")
        {
            return ellipseCodeItemsList.Select(item => item.code + separator + item.description).ToList();
        }
    }

    public class Encryption : SharedClassLibrary.Utilities.Encryption
    {

    }

    public class FileWriter : SharedClassLibrary.Utilities.FileWriter
    {

    }

    public class IxConstantInteger : SharedClassLibrary.Utilities.IxConstantInteger
    {

    }

    public class InputBox : SharedClassLibrary.Utilities.InputBox
    {
    }

    public class MyKeyValuePair<TKey, TVal> : SharedClassLibrary.Utilities.MyKeyValuePair<TKey, TVal>
    {
        public MyKeyValuePair() : base() { }

        public MyKeyValuePair(TKey key, TVal val) : base(key, val)
        {
        }
    }

    public class Math : SharedClassLibrary.Utilities.Math
    {

    }
}