using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Windows.Forms;
using System.Web.Services.Ellipse;

namespace EllipseTemplateExcelAddIn
{
    [SuppressMessage("ReSharper", "AccessToStaticMemberViaDerivedType")]
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Excel.Application _excelApp;

        private const string SheetName01 = "TemplateSheet";
        private const string TableName01 = "TemplateTable";
        private const int TitleRow01 = 7;
        private const int ResultColumn01 = 4;

        private const string SheetName02 = "TemplateSheet2";
        private const string TableName02 = "TemplateTable2";
        private const int TitleRow02 = 7;
        private const int ResultColumn02 = 4;

        private const string ValidationSheetName = "ValidationSheet";
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new FormAuthenticate();
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }

            var defaultConfig = new Settings.Options();
            //defaultConfig.SetOption("OptionName1", "OptionValue1");
            //defaultConfig.SetOption("OptionName2", "OptionValue2");
            //defaultConfig.SetOption("OptionName3", "OptionValue3");

            var options = settings.GetOptionsSettings(defaultConfig);

            //Setting of Configuration Options from Config File (or default)
            //var optionItem1Value = MyUtilities.IsTrue(options.GetOptionValue("OptionName1"));
            //var optionItem1Value = options.GetOptionValue("OptionName2");
            //var optionItem1Value = options.GetOptionValue("OptionName3");

            //optionItem1.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.UpdateOptionsSettings(options);
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatMethod();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(@"El separador decimal configurado actualmente no es el punto. Se recomienda ajustar antes esta configuración para evitar que se ingresen valores que no corresponden con los del sistema Ellipse", @"ADVERTENCIA");
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            const string developerName1 = "Héctor Hernandez <hernandezrhectorj@gmail.com>";
            const string developerName2 = "Hugo Mendoza <huancone@gmail.com>";

            new AboutBoxExcelAddIn(developerName1, developerName2).ShowDialog();
        }

        private void btnStop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }

        private void btnExecute_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(() => ExecutionMethod());

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ExecutionMethod()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void FormatMethod()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.SetCursorWait();

                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                #region CONSTRUYO LA HOJA 1
                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;
                var tableName = TableName01;
                var sheetName = SheetName01;

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "TEMPLATE ADDIN - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                var districtList = Districts.GetDistrictList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 2, false);
                _cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);

                //
                _cells.GetCell(1, titleRow).Value = "COLUMNA1";
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, titleRow).Value = "COLUMNA2";
                _cells.GetCell(2, titleRow).AddComment("Comentario de Encabezado de Columna");
                _cells.GetCell(3, titleRow).Value = "COLUMNA3";
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleInformation;

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 2

                titleRow = TitleRow02;
                resultColumn = ResultColumn02;
                tableName = TableName02;
                sheetName = SheetName02;

                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "TEMPLATE ADDIN HOJA 2- ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                districtList = Districts.GetDistrictList();
                workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.SetValidationList(_cells.GetCell("B4"), ValidationSheetName, 2, false);
                _cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);

                //
                _cells.GetCell(1, titleRow).Value = "COLUMNA1";
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, titleRow).Value = "COLUMNA2";
                _cells.GetCell(2, titleRow).AddComment("Comentario de Encabezado de Columna");
                _cells.GetCell(3, titleRow).Value = "COLUMNA3";
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleInformation;

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatMehod()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }


        private void ExecutionMethod()
        {
            //if (_cells == null)
            //    _cells = new ExcelStyleCells(_excelApp);
            //_cells.SetCursorWait();

            //var tableName = TableName01;
            //var titleRow = TitleRow01;
            //var resultColumn = ResultColumn01;

            //_cells.ClearTableRangeColumn(tableName, resultColumn);

            //_eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            ////Creación del Servicio
            //var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            //var service = new NameService.NameService();
            //service.Url = urlService + "/NameService";
            
            ////Instanciar el Contexto de Operación
            //var opContext = new NameService.OperationContext
            //{
            //    district = _frmAuth.EllipseDsct,
            //    position = _frmAuth.EllipsePost,
            //    maxInstances = 100,
            //    maxInstancesSpecified = true,
            //    returnWarnings = Debugger.DebugWarnings,
            //    returnWarningsSpecified = true
            //};


            ////Instanciar el SOAP
            //ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //var i = titleRow + 1;
            //var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDsct) ?? "ICOR";
            //while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)))
            //{
            //    try
            //    {
            //        var column1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
            //        var column2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
            //        var column3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);


            //        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //        //Se cargan los parámetros de  la solicitud
            //        var request = new NameServiceCreateRequestDTO();
            //        request.column1 = column1;
            //        request.column2 = column2;
            //        request.column3 = column3;

            //        //se envía la acción
            //        var reply = service.action(opContext, request);

            //        //se analiza la respuesta y se hacen las acciones pertinentes
            //        var replyValue = reply.Value;

            //        //
            //        _cells.GetCell(resultColumn, i).Value = "REALIZADO " + replyValue;
            //        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
            //    }
            //    catch (Exception ex)
            //    {
            //        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
            //        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
            //        Debugger.LogError("RibbonEllipse.cs:ExecutionMethod()", ex.Message);
            //    }
            //    finally
            //    {
            //        _cells.GetCell(resultColumn, i).Select();
            //        i++;
            //    }
            //}
            //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            //if (_cells != null) _cells.SetCursorDefault();
        }
    }
}
