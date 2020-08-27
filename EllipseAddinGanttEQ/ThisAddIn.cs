using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using System.Windows.Forms;

namespace EllipseAddinGanttEQ
{
    public partial class ThisAddIn
    {
        //private Excel.Application _excelApp;
        //public event Microsoft.Office.Interop.Excel.WorkbookEvents_SheetChangeEventHandler SheetChange;
        //public event Microsoft.Office.Interop.Excel.WorkbookEvents_SheetChangeEventHandler SheetChange;
        //internal Microsoft.Office.Interop.Excel.WorkbookEvents_SheetChangeEventHandler SheetChange;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //_excelApp = Globals.ThisAddIn.Application;
            /*_excelApp.Visible = true;
            _excelApp.ScreenUpdating = true;
            _excelApp.DisplayAlerts = true;
            */
            //_excelApp.EnableEvents = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            /*_excelApp = Globals.ThisAddIn.Application;
            _excelApp.Visible = true;
            //_excelApp.Visible = false;
            _excelApp.ScreenUpdating = true;
            _excelApp.DisplayAlerts = true;
            */
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}