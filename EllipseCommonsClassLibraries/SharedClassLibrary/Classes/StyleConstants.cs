using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

//Shared Class Library - ExcelStyleCells
//Desarrollado por:
//Héctor J Hernández R <hernandezrhectorj@gmail.com>
//Hugo A Mendoza B <hugo.mendoza@hambings.com.co>

namespace SharedClassLibrary.Classes
{
    /// <summary>
    /// Estilos Predeterminados del Sistemas
    /// </summary>
    public class StyleConstants
    {
        //PASOS PARA AGREGAR UN ESTILO A LA CLASE
        //1. Agregar la variable de constante en esta clase
        //2. Adicionarla a la lista del método GetStyleName
        //3. Adicionarla a ExcelStyleCells.CreateStyle()

        public const string Normal = "MyNormal";
        public const string Success = "Success";
        public const string Warning = "Warning";
        public const string Error = "Error";
        public const string HeaderDefault = "HeaderDefault";
        public const string HeaderSize17 = "HeaderSize17";
        public const string TitleDefault = "TitleDefault";
        public const string TitleRequired = "TitleRequired";
        public const string TitleOptional = "TitleOptional";
        public const string TitleInformation = "TitleInformation";
        public const string TitleAction = "TitleAction";
        public const string TitleAdditional = "TitleAdditional";
        public const string TitleResult = "TitleResult";
        public const string Option = "Option";
        public const string Select = "Select";
        public const string Disabled = "Disabled";
        public const string Time = "Time";
        public const string ItalicSmall = "ItalicSmall";
        public static List<string> GetStyleListName()
        {
            var styleConstantsList = new List<string>
            {
                Normal,
                Success,
                Warning,
                Error,
                HeaderDefault,
                HeaderSize17,
                TitleDefault,
                TitleRequired,
                TitleOptional,
                TitleInformation,
                TitleAction,
                TitleAdditional,
                TitleResult,
                Option,
                Select,
                Disabled,
                Time,
                ItalicSmall
            };

            return styleConstantsList;
        }

        public static class TableStyleConstants
        {
            public const string DefaultTableStyle = "TableStyleLight8";
        }
    }
    //Formatos de Número para Celdas del Sistema
    public class NumberFormatConstants
    {
        public const string General = "General";
        public const string Text = "@";
        public const string Integer = "0";
        public const string Number = "0";
        public const string Date = "yyyyMMdd";
        public const string DateTime = "yyyyMMdd HHmmss";
        public const string Percentage = "###,##%";
    }

    public static class LanguageSettingConstants
    {
        public static readonly string ListSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
        public static readonly string DecimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

    }
}
