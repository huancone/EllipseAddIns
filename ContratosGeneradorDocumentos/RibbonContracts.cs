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
        private const int TitleRow01 = 8;
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
            var fileUrl = "" + _cells.GetCell(2, 4).Value;
            var fileName = "" + _cells.GetCell(3, 4).Value;

            var destUrl = "" + _cells.GetCell(2, 5).Value;
            var destName = "" + _cells.GetCell(3, 5).Value;

            var list = new List<KeyValuePair<string, string>>();
            for (int i = 1; i <= 13; i++)
            {
                list.Add(new KeyValuePair<string, string>(_cells.GetCell(2, TitleRow01 + i).Value, "" + _cells.GetCell(3, TitleRow01 + i).Value));
            }
            
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

                _cells.GetCell("A3").Value = "TIPO";
                _cells.GetCell("B3").Value = "RUTA";
                _cells.GetCell("C3").Value = "NOMBRE ARCHIVO";
                _cells.GetCell("A4").Value = "Plantilla";
                _cells.GetCell("B4").Value = @"c:\ellipse";
                _cells.GetCell("C4").Value = @"Plantilla.docx";
                _cells.GetCell("A5").Value = "Destino";
                _cells.GetCell("B5").Value = @"c:\ellipse";
                _cells.GetCell("C5").Value = @"Resultado";

                _cells.GetRange("A3", "C3").Style = StyleConstants.TitleOptional;
                _cells.GetRange("A4", "A5").Style = StyleConstants.Option;
                _cells.GetRange("B4", "C5").Style = StyleConstants.Select;


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

                _cells.GetCell(1, titleRow + 1).Value = "Contrato";
                _cells.GetCell(2, titleRow + 1).Value = "nroContrato";
                _cells.GetCell(3, titleRow + 1).Value = "123456789";

                _cells.GetCell(1, titleRow + 2).Value = "Tipo de Anexo";
                _cells.GetCell(2, titleRow + 2).Value = "tituloAnexo";
                _cells.GetCell(3, titleRow + 2).Value = "ALCANCE DE LOS SERVICIOS";

                _cells.GetCell(1, titleRow + 3).Value = "Objeto del Contrato";
                _cells.GetCell(2, titleRow + 3).Value = "objetoContrato";
                _cells.GetCell(3, titleRow + 3).Value = "SOPORTE EN DESARROLLO SOFTWARE PARE LA GENERACIÓN AUTOMÁTICA DE DOCUMENTOS";

                _cells.GetCell(1, titleRow + 4).Value = "Duración (Número)";
                _cells.GetCell(2, titleRow + 4).Value = "duracionNumero";
                _cells.GetCell(3, titleRow + 4).Value = "4";

                _cells.GetCell(1, titleRow + 5).Value = "Duración (Letras)";
                _cells.GetCell(2, titleRow + 5).Value = "duracionLetras";
                _cells.GetCell(3, titleRow + 5).Value = "cuatro";

                _cells.GetCell(1, titleRow + 6).Value = "Unidades (Días, Semanas, Meses, Años)";
                _cells.GetCell(2, titleRow + 6).Value = "duracionUnidades";
                _cells.GetCell(3, titleRow + 6).Value = "semanas";

                var unitsList = new List<string> {"Días", "Semanas", "Meses", "Años"};
                _cells.SetValidationList(_cells.GetCell(3, titleRow + 6), unitsList);


                _cells.GetCell(1, titleRow + 7).Value = "Prórroga (Número)";
                _cells.GetCell(2, titleRow + 7).Value = "duracionProrrogaNumero";
                _cells.GetCell(3, titleRow + 7).Value = "2";

                _cells.GetCell(1, titleRow + 8).Value = "Descripción del Servicio";
                _cells.GetCell(2, titleRow + 8).Value = "descripcionServicios";
                _cells.GetCell(3, titleRow + 8).Value = "Soporte en el desarrollo de una aplicación para la generación automática de documentos para los anexos de contratos";

                _cells.GetCell(1, titleRow + 9).Value = "Lugar";
                _cells.GetCell(2, titleRow + 9).Value = "lugarContrato";
                _cells.GetCell(3, titleRow + 9).Value = "";

                var locList = new List<string> { "LA MINA", "BARRANQUILLA", "BOGOTÁ", "LA MINA Y PUERTO BOLÍVAR", "PUERTO BOLÍVAR" };
                _cells.SetValidationList(_cells.GetCell(3, titleRow + 9), locList);

                _cells.GetCell(1, titleRow + 10).Value = "Perfil Profesional 1";
                _cells.GetCell(2, titleRow + 10).Value = "perfilProfesional:1";
                _cells.GetCell(3, titleRow + 10).Value = "Profesional";

                _cells.GetCell(1, titleRow + 11).Value = "Descripción Perfil 1";
                _cells.GetCell(2, titleRow + 11).Value = "perfilProfesionalDesc:1";
                _cells.GetCell(3, titleRow + 11).Value = @"Ingeniero de sistemas, con competencias y demostrados conocimientos en las herramientas relacionadas y experiencia mínima de cinco(5) años en manejo de sistemas de información de mantenimiento.";

                _cells.GetCell(1, titleRow + 12).Value = "Perfil Profesional 2";
                _cells.GetCell(2, titleRow + 12).Value = "perfilProfesional:2";
                _cells.GetCell(3, titleRow + 12).Value = "Desarrollador";

                _cells.GetCell(1, titleRow + 13).Value = "Descripción Perfil 2";
                _cells.GetCell(2, titleRow + 13).Value = "perfilProfesionalDesc:2";
                _cells.GetCell(3, titleRow + 13).Value = "Desarrollador software con competencias en c#, vsto, arquitectura de software y certificación de base de datos.";


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
