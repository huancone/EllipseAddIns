using System;
using Office = Microsoft.Office.Core;

namespace EllipseMSO627RodillosAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region Código generado por VSTO

        /// <summary>
        ///     Método necesario para admitir el Diseñador. No se puede modificar
        ///     el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}