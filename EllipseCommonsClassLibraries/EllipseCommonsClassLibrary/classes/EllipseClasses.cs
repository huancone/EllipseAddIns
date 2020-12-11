using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SharedClassLibrary.Classes;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseCommonsClassLibrary.Classes
{
    public class ExcelStyleCells : SharedClassLibrary.Vsto.Excel.ExcelStyleCells
    {
        public ExcelStyleCells(Application excelApp, bool alwaysActiveSheet = true) : base(excelApp, alwaysActiveSheet)
        {

        }

        public ExcelStyleCells(Application excelApp, string sheetName) : base(excelApp, sheetName)
        {

        }
    }

    public class StyleConstants : SharedClassLibrary.Vsto.Excel.StyleConstants
    {

    }

    public class NumberFormatConstants : SharedClassLibrary.Vsto.Excel.NumberFormatConstants
    {

    }

    public class ReplyMessage : SharedClassLibrary.Classes.ReplyMessage
    {

    }
}
