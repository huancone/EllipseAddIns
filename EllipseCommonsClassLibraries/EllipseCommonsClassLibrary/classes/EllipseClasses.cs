using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommonsClassLibrary.Classes;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseCommonsClassLibrary.Classes
{
    public class ExcelStyleCells : CommonsClassLibrary.Classes.ExcelStyleCells
    {
        public ExcelStyleCells(Application excelApp, bool alwaysActiveSheet = true) : base(excelApp, alwaysActiveSheet)
        {

        }

        public ExcelStyleCells(Application excelApp, string sheetName) : base(excelApp, sheetName)
        {

        }
    }

    public class StyleConstants : CommonsClassLibrary.Classes.StyleConstants
    {

    }

    public class NumberFormatConstants : CommonsClassLibrary.Classes.NumberFormatConstants
    {

    }

    public class ReplyMessage : CommonsClassLibrary.Classes.ReplyMessage
    {

    }
}
