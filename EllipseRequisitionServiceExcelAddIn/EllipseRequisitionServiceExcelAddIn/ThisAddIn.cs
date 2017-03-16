using System;
using Office = Microsoft.Office.Core;

namespace EllipseRequisitionServiceExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
