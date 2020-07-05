using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using System.Collections;
using DocumentGeneratorClassLibrary;
using EllipseCommonsClassLibrary.Settings;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ContratosGeneradorDocumentos
{
    public partial class RibbonContracts
    {
        private ExcelStyleCells _cells;
        private Application _excelApp;
        private Thread _thread;
        private CommonSettings _settings;

        //Hojas
        private const string ValidationSheetName = "Validacion";
        private const string SheetName01 = "Planeados";
        private const string TableName01 = "JobResources";
        private const int TitleRow01 = 7;
        private const int ResultColumn01 = 4;


        private void RibbonContracts_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            LoadSettings();
        }

        private void LoadSettings()
        {

        }
        private void btnAction_Click(object sender, RibbonControlEventArgs e)
        {
            var fileUrl = "" + _cells.GetCell(2, 3).Value;
            var fileName = "" + _cells.GetCell(4, 3).Value;

            var destUrl = "" + _cells.GetCell(2, 4).Value;
            var destName = "" + _cells.GetCell(4, 4).Value;

            var list = new List<KeyValuePair<string, string>>();
            list.Add(new KeyValuePair<string, string>("firstName", "" + _cells.GetCell(3, TitleRow01 + 1).Value));
            list.Add(new KeyValuePair<string, string>("surName", "" + _cells.GetCell(3, TitleRow01 + 2).Value));
            list.Add(new KeyValuePair<string, string>("address", "" + _cells.GetCell(3, TitleRow01 + 3).Value));
            list.Add(new KeyValuePair<string, string>("city", "" + _cells.GetCell(3, TitleRow01 + 4).Value));
            list.Add(new KeyValuePair<string, string>("zipCode", "" + _cells.GetCell(3, TitleRow01 + 5).Value));
            list.Add(new KeyValuePair<string, string>("phone", "" + _cells.GetCell(3, TitleRow01 + 6).Value));
            list.Add(new KeyValuePair<string, string>("email", "" + _cells.GetCell(3, TitleRow01 + 7).Value));
            list.Add(new KeyValuePair<string, string>("webSite", @"" + _cells.GetCell(3, TitleRow01 + 8).Value));
            
            Finder.Execute(fileUrl, fileName, destUrl, destName, list);

        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void FormatSheet()
        {

            try
            {
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 2)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación
                _cells.SetCursorWait();

                var sheetName = SheetName01;
                var titleRow = TitleRow01;
                var tableName = TableName01;
                var resultColumn = ResultColumn01;

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("B1").Value = "CONTRATOS - GENERADOR DE DOCUMENTOS";

                _cells.GetRange("A1", "B1").Style = StyleConstants.HeaderDefault;
                _cells.GetRange("B1", "D1").Merge();

                _cells.GetCell("A3").Value = "Ruta Original";
                _cells.GetCell("B3").Value = @"c:\ellipse";
                _cells.GetCell("A4").Value = "Ruta Destino";
                _cells.GetCell("B4").Value = @"c:\ellipse";
                _cells.GetCell("C3").Value = "Nombre Original";
                _cells.GetCell("D3").Value = @"prueba.docx";
                _cells.GetCell("C4").Value = "Nombre Destino";
                _cells.GetCell("D4").Value = @"resultado";

                _cells.GetCell(1, titleRow).Value = "Descripción";
                _cells.GetCell(2, titleRow).Value = "Etiqueta";
                _cells.GetCell(3, titleRow).Value = "Valor";
                _cells.GetCell(resultColumn, titleRow).Value = "Result";

                #region Styles
                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleOptional;
                #endregion

                #region Instructions

                _cells.GetCell("E1").Value = "OBLIGATORIO";
                _cells.GetCell("E1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("E2").Value = "OPCIONAL";
                _cells.GetCell("E2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("E3").Value = "INFORMATIVO";
                _cells.GetCell("E3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("E4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("E4").Style = StyleConstants.TitleAction;
                _cells.GetCell("E5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("E5").Style = StyleConstants.TitleAdditional;

                #endregion

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);

                _cells.GetCell(1, titleRow + 1).Value = "Nombre";
                _cells.GetCell(2, titleRow + 1).Value = "firstName";
                _cells.GetCell(3, titleRow + 1).Value = "Jessica";
                _cells.GetCell(1, titleRow + 2).Value = "Apellido";
                _cells.GetCell(2, titleRow + 2).Value = "surName";
                _cells.GetCell(3, titleRow + 2).Value = "Videz";
                _cells.GetCell(1, titleRow + 3).Value = "Dirección";
                _cells.GetCell(2, titleRow + 3).Value = "address";
                _cells.GetCell(3, titleRow + 3).Value = "Casita de Dios";
                _cells.GetCell(1, titleRow + 4).Value = "Ciudad";
                _cells.GetCell(2, titleRow + 4).Value = "city";
                _cells.GetCell(3, titleRow + 4).Value = "Medellín";
                _cells.GetCell(1, titleRow + 5).Value = "Código Postal";
                _cells.GetCell(2, titleRow + 5).Value = "zipCode";
                _cells.GetCell(3, titleRow + 5).Value = "0012345";
                _cells.GetCell(1, titleRow + 6).Value = "Teléfono";
                _cells.GetCell(2, titleRow + 6).Value = "phone";
                _cells.GetCell(3, titleRow + 6).Value = "555 77 88";
                _cells.GetCell(1, titleRow + 7).Value = "Correo";
                _cells.GetCell(2, titleRow + 7).Value = "email";
                _cells.GetCell(3, titleRow + 7).Value = "mamasita@bella.com";
                _cells.GetCell(1, titleRow + 8).Value = "Sitio";
                _cells.GetCell(2, titleRow + 8).Value = "webSite";
                _cells.GetCell(3, titleRow + 8).Value = @"http://www.dejaelchisme.com";


                ((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Cells.Columns.AutoFit();

                ((Worksheet)_excelApp.ActiveWorkbook.Sheets[1]).Select(Type.Missing);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

    }
}
