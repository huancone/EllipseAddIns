using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseWorkRequestClassLibrary;
using EllipseWorkRequestClassLibrary.WorkRequestService;
using EllipseReferenceCodesClassLibrary;
using EllipseStdTextClassLibrary;

namespace EllipseWorkRequestExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        EllipseFunctions _eFunctions = new EllipseFunctions();
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;
        private const string SheetName01 = "WorkRequest";
        private const string SheetName02 = "WorkRequestClose";
        private const string SheetName03 = "WorkRequestsReferences";
        private const string SheetNameM01 = "WorkRequestMntto";
        private const string SheetNameM02 = "WorkRequestMnttoClose";
        private const string SheetNameM03 = "WorkRequestsMnttoSLA";
        //private const string SheetName04 = "WorkOrdersRelated";
        private const int TitleRow01 = 9;
        private const int TitleRow02 = 6;
        private const int TitleRow03 = 9;
        private const int ResultColumn01 = 38;
        private const int ResultColumn02 = 5;
        private const int ResultColumn03 = 23;
        private const int ResultColumnM01 = 46;
        private const int ResultColumnM02 = 5;
        private const int ResultColumnM03 = 14;
        private const string TableName01 = "WorkRequestTable";
        private const string TableName02 = "WorkRequestCloseTable";
        private const string TableName03 = "WorkRequestsReferencesTable";
        private const string TableNameM01 = "WorkRequestTable";
        private const string TableNameM02 = "WorkRequestCloseTable";
        private const string TableNameM03 = "WorkRequestSLATable";
        //private const string TableName04 = "WorkOrdersRelatedTable";
        private const string ValidationSheetName = "ValidationSheet";

        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            
            _eFunctions.DebugQueries = false;
            _eFunctions.DebugErrors = false;
            _eFunctions.DebugWarnings = false;
            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }   

        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }
        private void btnFormatMantto_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetMtto();
        }
        private void btnReviewWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameM01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReReviewWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameM01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReReviewWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCreateWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(CreateWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameM01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(CreateWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnModifyWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ModifyWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ModifyWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnCloseWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(CloseWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CloseWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReOpenWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReOpenWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DeleteWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnDeleteWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(DeleteWorkRequestList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DeleteWorkRequest()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnSetSla_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(SetSlaList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:SetSlaList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        
        private void btnResetSla_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ResetSlaList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:SetSlaList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReviewRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewReferenceCodesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewReferenceCodesList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReReviewRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewReferenceCodesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReReviewReferenceCodesList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateReferenceCodesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:UpdateReferenceCodesList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            
        }

        private void UpdateReferenceCodesList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                throw new NotImplementedException();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:UpdateReferenceCodesList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace,
                    _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
            
        }


        private void btnCleanSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
            _cells.ClearTableRange(TableName02);
            _cells.ClearTableRange(TableName03);
        }
        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }
        
        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK REQUEST - ELLIPSE 8";
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

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.GetCell("A5").Value = "STATUS";
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("D4").AddComment("YYYYMMDD");

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell("B4"), GroupConstants.GetWorkGroupList().Select(g => g.Name).ToList(), ValidationSheetName, 2, false);
                var wrStatusList = WrStatusList.GetStatusNames();
                wrStatusList.Add(WrStatusList.Uncompleted);
                _cells.SetValidationList(_cells.GetCell("B5"), wrStatusList, ValidationSheetName, 3, false);

                //Valores predeterminados
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);


                _cells.GetRange(2, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(1, TitleRow01, ResultColumn01 - 1, TitleRow01).Style = StyleConstants.TitleRequired;
                for (var i = 2; i < ResultColumn01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                //GENERAL
                _cells.GetCell(02, TitleRow01 - 2).Value = "GENERAL";
                _cells.MergeCells(02, TitleRow01 - 2, 08, TitleRow01 - 2);

                _cells.GetCell(01, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(02, TitleRow01).Value = "REQUEST ID";
                _cells.GetCell(03, TitleRow01).Value = "WR STATUS";
                _cells.GetCell(03, TitleRow01 - 1).Value2 = "";
                _cells.GetCell(03, TitleRow01 - 1).ClearComments();
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(04, TitleRow01).Value = "DESCRIPTION 1";
                _cells.GetCell(05, TitleRow01).Value = "DESCRIPTION 2";
                _cells.GetCell(05, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(06, TitleRow01).Value = "EQUIPMENT";

                //WORK
                _cells.GetCell(07, TitleRow01 - 2).Value = "WORK";
                _cells.MergeCells(07, TitleRow01 - 2, 12, TitleRow01 - 2);

                _cells.GetCell(07, TitleRow01).Value = "EMPLOYEE";
                _cells.GetCell(08, TitleRow01).Value = "CLASSIFICATION";
                _cells.GetCell(09, TitleRow01).Value = "REQUEST TYPE";
                _cells.GetCell(10, TitleRow01).Value = "USER STATUS";
                _cells.GetCell(10, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, TitleRow01).Value = "PRIORITY";
                _cells.GetCell(12, TitleRow01).Value = "REGION";
                _cells.GetCell(12, TitleRow01).Style = StyleConstants.TitleOptional;

                var classificationItemCodeList = _eFunctions.GetItemCodes("RQCL");
                var classificationList = Utils.GetCodeList(classificationItemCodeList);
                _cells.SetValidationList(_cells.GetCell(08, TitleRow01 + 1), classificationList, ValidationSheetName, 4, false);

                var reqTypeItemCodeList = WoTypeMtType.GetWoTypeList();
                var requestTypeList = Utils.GetCodeList(reqTypeItemCodeList);
                _cells.SetValidationList(_cells.GetCell(09, TitleRow01 + 1), requestTypeList, ValidationSheetName, 5, false);

                var usTypeCodeList = Utils.GetCodeList(WorkRequestActions.GetUserStatusCodeList(_eFunctions));
                _cells.SetValidationList(_cells.GetCell(10, TitleRow01 + 1), usTypeCodeList, ValidationSheetName, 6, false);
                
                var priorityList = Utils.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                _cells.SetValidationList(_cells.GetCell(11, TitleRow01 + 1), priorityList, ValidationSheetName, 7, false);

                var regionItemCodeList = _eFunctions.GetItemCodes("REGN");
                var regionList = Utils.GetCodeList(regionItemCodeList);
                _cells.SetValidationList(_cells.GetCell(12, TitleRow01 + 1), regionList, ValidationSheetName, 8, false);
                
                //SOURCE
                _cells.GetCell(13, TitleRow01 - 2).Value = "SOURCE";
                _cells.MergeCells(13, TitleRow01 - 2, 15, TitleRow01 - 2);

                _cells.GetCell(13, TitleRow01).Value = "CONTACT ID";
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(14, TitleRow01).Value = "REQ.SOURCE";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(15, TitleRow01).Value = "S.REFERENCE";
                _cells.GetCell(15, TitleRow01).Style = StyleConstants.TitleOptional;

                var reqSourceItemCodeList = _eFunctions.GetItemCodes("RQSC");
                var requestSourceList = Utils.GetCodeList(reqSourceItemCodeList);
                _cells.SetValidationList(_cells.GetCell(14, TitleRow01 + 1), requestSourceList, ValidationSheetName, 9, false);

                //DATES
                _cells.GetCell(16, TitleRow01 - 2).Value = "DATES";
                _cells.MergeCells(16, TitleRow01 - 2, 22, TitleRow01 - 2);

                _cells.GetCell(16, TitleRow01).Value = "REQ DATE";
                _cells.GetCell(16, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(17, TitleRow01).Value = "REQ TIME";
                _cells.GetCell(17, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(18, TitleRow01).Value = "RAISED BY";
                _cells.GetCell(19, TitleRow01).Value = "RAISED DATE";
                _cells.GetCell(20, TitleRow01).Value = "RAISED TIME";
                _cells.GetCell(21, TitleRow01).Value = "COMPLETED BY";
                _cells.GetCell(22, TitleRow01).Value = "CLOSED DATE";
                _cells.GetRange(18, TitleRow01 - 1, 22, TitleRow01 - 1).Value2 = "";
                _cells.GetRange(18, TitleRow01 - 1, 22, TitleRow01 - 1).ClearComments();
                _cells.GetRange(18, TitleRow01, 22, TitleRow01).Style = StyleConstants.TitleInformation;

                //ASSIGN
                _cells.GetCell(23, TitleRow01 - 2).Value = "ASSIGN";
                _cells.MergeCells(23, TitleRow01 - 2, 24, TitleRow01 - 2);
                _cells.GetCell(23, TitleRow01).Value = "ASSIGN TO";
                _cells.GetCell(24, TitleRow01).Value = "OWNER ID";


                //ESTIMATE
                _cells.GetCell(25, TitleRow01 - 2).Value = "ESTIMATE";
                _cells.MergeCells(25, TitleRow01 - 2, 27, TitleRow01 - 2);

                _cells.GetCell(25, TitleRow01).Value = "ESTIMATE NO";
                _cells.GetCell(25, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(26, TitleRow01).Value = "STD JOB NO";
                _cells.GetCell(26, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(27, TitleRow01).Value = "SJ DISTRICT";
                _cells.GetCell(27, TitleRow01).Style = StyleConstants.TitleOptional;
                //SERVICE LEVEL AGREEMENT
                _cells.GetCell(28, TitleRow01 - 2).Value = "SERVICE LEVEL AGREEMENT";
                _cells.GetCell(28, TitleRow01 - 2).AddComment("Esta sección solo se actualiza con las acciones SLA");
                _cells.MergeCells(28, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2);


                _cells.GetCell(28, TitleRow01).Value = "SL_AGREEMENT";
                _cells.GetCell(29, TitleRow01).Value = "SLA_FAILURE_CODE";
                _cells.GetCell(30, TitleRow01).Value = "SLA_START_DATE";
                _cells.GetCell(31, TitleRow01).Value = "SLA_START_TIME";
                _cells.GetCell(32, TitleRow01).Value = "SLA_DUE_DATE";
                _cells.GetCell(33, TitleRow01).Value = "SLA_DUE_TIME";
                _cells.GetCell(34, TitleRow01).Value = "SLA_DUE_DAYS";
                _cells.GetCell(35, TitleRow01).Value = "SLA_WARN_DATE";
                _cells.GetCell(36, TitleRow01).Value = "SLA_WARN_TIME";
                _cells.GetCell(37, TitleRow01).Value = "SLA_WARN_DAYS";
                //
                _cells.GetRange(28, TitleRow01, ResultColumn01-1, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;

                
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                ////CONSTRUYO LA HOJA 2 - CLOSE WR
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;
                
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "CLOSE WORK REQUEST - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = StyleConstants.TitleAdditional;

                _cells.GetRange(1, TitleRow02, ResultColumn02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow02).Value = "REQUEST ID";
                _cells.GetCell(2, TitleRow02).Value = "CLOSED BY";
                _cells.GetCell(3, TitleRow02).Value = "CLOSED DATE";
                _cells.GetCell(4, TitleRow02).Value = "CLOSED TIME";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                ////CONSTRUYO LA HOJA 3 RERFERENCE CODES WR
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "REFERENCE CODES WORK REQUEST - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = StyleConstants.TitleAdditional;

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.GetCell("A5").Value = "STATUS";
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("D4").AddComment("YYYYMMDD");

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell("B4"), ValidationSheetName, 2, false);
                _cells.SetValidationList(_cells.GetCell("B5"), ValidationSheetName, 3, false);
                //Valores predeterminados
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

                _cells.GetRange(1, TitleRow03, ResultColumn03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;
                for (var i = 4; i < ResultColumn03; i++)
                {
                    _cells.GetCell(i, TitleRow03 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow03 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow03 - 1).Value = "true";
                }
                _cells.GetRange(1, TitleRow03, ResultColumn03, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(02, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(03, TitleRow03).Style = StyleConstants.TitleInformation;

                _cells.GetCell(01, TitleRow03).Value = "WORKGROUP"; 
                _cells.GetCell(02, TitleRow03).Value = "REQUEST ID";
                _cells.GetCell(03, TitleRow03).Value = "DESCRIPTION";
                _cells.GetCell(04, TitleRow03).Value = "DESC EXTEND HEADER";
                _cells.GetCell(05, TitleRow03).Value = "DESC EXTEND BODY";
                _cells.GetCell(06, TitleRow03).Value = "OT ORIGEN";
                _cells.GetCell(07, TitleRow03).Value = "STOCK CODE 1";
                _cells.GetCell(08, TitleRow03).Value = "SC QTY 1";
                _cells.GetCell(09, TitleRow03).Value = "STOCK CODE 2";
                _cells.GetCell(10, TitleRow03).Value = "SC QTY 2";
                _cells.GetCell(11, TitleRow03).Value = "STOCK CODE 3";
                _cells.GetCell(12, TitleRow03).Value = "SC QTY 3";
                _cells.GetCell(13, TitleRow03).Value = "STOCK CODE 4";
                _cells.GetCell(14, TitleRow03).Value = "SC QTY 4";
                _cells.GetCell(15, TitleRow03).Value = "STOCK CODE 5";
                _cells.GetCell(16, TitleRow03).Value = "SC QTY 5";
                _cells.GetCell(17, TitleRow03).Value = "H.HOMBRE RES";
                _cells.GetCell(18, TitleRow03).Value = "H.HOMBRE QTY";
                _cells.GetCell(19, TitleRow03).Value = "DURACIÓN TAREA";
                _cells.GetCell(20, TitleRow03).Value = "EQUIPO DETENIDO";
                _cells.GetCell(21, TitleRow03).Value = "RAISED REP.";
                _cells.GetCell(22, TitleRow03).Value = "CAMBIO HORA";
                _cells.GetCell(22, TitleRow03).AddComment("HH:MM");
                _cells.GetCell(ResultColumn03, TitleRow03).Value = "RESULTADO";
                _cells.GetCell(ResultColumn03, TitleRow03).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow03 + 1, ResultColumn03, TitleRow03 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumn03, TitleRow03 + 1), TableName03);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                
                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void FormatSheetMtto()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameM01;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK REQUEST - ELLIPSE 8";
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

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.GetCell("A5").Value = "STATUS";
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("D4").AddComment("YYYYMMDD");

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell("B4"), GroupConstants.GetWorkGroupList().Select(g => g.Name).ToList(), ValidationSheetName, 2, false);
                var wrStatusList = WrStatusList.GetStatusNames();
                wrStatusList.Add(WrStatusList.Uncompleted);
                _cells.SetValidationList(_cells.GetCell("B5"), wrStatusList, ValidationSheetName, 3, false);

                //Valores predeterminados
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);


                _cells.GetRange(2, TitleRow01 - 2, ResultColumnM01 - 1, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(1, TitleRow01, ResultColumnM01 - 1, TitleRow01).Style = StyleConstants.TitleRequired;
                for (var i = 2; i < ResultColumnM01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }
                _cells.GetRange(1, TitleRow01 + 1, ResultColumnM01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                //GENERAL
                _cells.GetCell(02, TitleRow01 - 2).Value = "GENERAL";
                _cells.MergeCells(02, TitleRow01 - 2, 08, TitleRow01 - 2);

                _cells.GetCell(01, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(02, TitleRow01).Value = "REQUEST ID";
                _cells.GetCell(03, TitleRow01).Value = "WR STATUS";
                _cells.GetCell(03, TitleRow01 - 1).Value2 = "";
                _cells.GetCell(03, TitleRow01 - 1).ClearComments();
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(04, TitleRow01).Value = "DESCRIPTION 1";
                _cells.GetCell(05, TitleRow01).Value = "DESCRIPTION 2";
                _cells.GetCell(05, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(06, TitleRow01).Value = "EQUIPMENT";

                //WORK
                _cells.GetCell(07, TitleRow01 - 2).Value = "WORK";
                _cells.MergeCells(07, TitleRow01 - 2, 12, TitleRow01 - 2);

                _cells.GetCell(07, TitleRow01).Value = "EMPLOYEE";
                _cells.GetCell(08, TitleRow01).Value = "CLASSIFICATION";
                _cells.GetCell(09, TitleRow01).Value = "REQUEST TYPE";
                _cells.GetCell(10, TitleRow01).Value = "USER STATUS";
                _cells.GetCell(10, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, TitleRow01).Value = "PRIORITY";
                _cells.GetCell(12, TitleRow01).Value = "REGION";
                _cells.GetCell(12, TitleRow01).Style = StyleConstants.TitleOptional;

                var classificationItemCodeList = _eFunctions.GetItemCodes("RQCL");
                var classificationList = Utils.GetCodeList(classificationItemCodeList);
                _cells.SetValidationList(_cells.GetCell(08, TitleRow01 + 1), classificationList, ValidationSheetName, 4, false);

                var reqTypeItemCodeList = WoTypeMtType.GetWoTypeList();
                var requestTypeList = Utils.GetCodeList(reqTypeItemCodeList);
                _cells.SetValidationList(_cells.GetCell(09, TitleRow01 + 1), requestTypeList, ValidationSheetName, 5, false);

                var usTypeCodeList = Utils.GetCodeList(WorkRequestActions.GetUserStatusCodeList(_eFunctions));
                _cells.SetValidationList(_cells.GetCell(10, TitleRow01 + 1), usTypeCodeList, ValidationSheetName, 6, false);

                var priorityList = Utils.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                _cells.SetValidationList(_cells.GetCell(11, TitleRow01 + 1), priorityList, ValidationSheetName, 7, false);

                var regionItemCodeList = _eFunctions.GetItemCodes("REGN");
                var regionList = Utils.GetCodeList(regionItemCodeList);
                _cells.SetValidationList(_cells.GetCell(12, TitleRow01 + 1), regionList, ValidationSheetName, 8, false);

                //SOURCE
                _cells.GetCell(13, TitleRow01 - 2).Value = "SOURCE";
                _cells.MergeCells(13, TitleRow01 - 2, 15, TitleRow01 - 2);

                _cells.GetCell(13, TitleRow01).Value = "CONTACT ID";
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(14, TitleRow01).Value = "REQ.SOURCE";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(15, TitleRow01).Value = "S.REFERENCE";
                _cells.GetCell(15, TitleRow01).Style = StyleConstants.TitleOptional;

                var reqSourceItemCodeList = _eFunctions.GetItemCodes("RQSC");
                var requestSourceList = Utils.GetCodeList(reqSourceItemCodeList);
                _cells.SetValidationList(_cells.GetCell(14, TitleRow01 + 1), requestSourceList, ValidationSheetName, 9, false);

                //DATES
                _cells.GetCell(16, TitleRow01 - 2).Value = "DATES";
                _cells.MergeCells(16, TitleRow01 - 2, 22, TitleRow01 - 2);

                _cells.GetCell(16, TitleRow01).Value = "REQ DATE";
                _cells.GetCell(16, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(17, TitleRow01).Value = "REQ TIME";
                _cells.GetCell(17, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(18, TitleRow01).Value = "RAISED BY";
                _cells.GetCell(19, TitleRow01).Value = "RAISED DATE";
                _cells.GetCell(20, TitleRow01).Value = "RAISED TIME";
                _cells.GetCell(21, TitleRow01).Value = "COMPLETED BY";
                _cells.GetCell(22, TitleRow01).Value = "CLOSED DATE";
                _cells.GetRange(18, TitleRow01 - 1, 22, TitleRow01 - 1).Value2 = "";
                _cells.GetRange(18, TitleRow01 - 1, 22, TitleRow01 - 1).ClearComments();
                _cells.GetRange(18, TitleRow01, 22, TitleRow01).Style = StyleConstants.TitleInformation;

                //ASSIGN
                _cells.GetCell(23, TitleRow01 - 2).Value = "ASSIGN";
                _cells.MergeCells(23, TitleRow01 - 2, 24, TitleRow01 - 2);
                _cells.GetCell(23, TitleRow01).Value = "ASSIGN TO";
                _cells.GetCell(24, TitleRow01).Value = "OWNER ID";


                //ESTIMATE
                _cells.GetCell(25, TitleRow01 - 2).Value = "ESTIMATE";
                _cells.MergeCells(25, TitleRow01 - 2, 27, TitleRow01 - 2);

                _cells.GetCell(25, TitleRow01).Value = "ESTIMATE NO";
                _cells.GetCell(25, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(26, TitleRow01).Value = "STD JOB NO";
                _cells.GetCell(26, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(27, TitleRow01).Value = "SJ DISTRICT";
                _cells.GetCell(27, TitleRow01).Style = StyleConstants.TitleOptional;
                //REFERENCES CODES
                _cells.GetCell(28, TitleRow01 - 2).Value = "REFERENCES CODE";
                _cells.MergeCells(28, TitleRow01 - 2, ResultColumnM01 - 1, TitleRow01 - 2);

                _cells.GetCell(28, TitleRow03).Value = "DESC EXTEND HEADER";
                _cells.GetCell(29, TitleRow03).Value = "DESC EXTEND BODY";
                _cells.GetCell(30, TitleRow03).Value = "OT ORIGEN";
                _cells.GetCell(31, TitleRow03).Value = "STOCK CODE 1";
                _cells.GetCell(32, TitleRow03).Value = "SC QTY 1";
                _cells.GetCell(33, TitleRow03).Value = "STOCK CODE 2";
                _cells.GetCell(34, TitleRow03).Value = "SC QTY 2";
                _cells.GetCell(35, TitleRow03).Value = "STOCK CODE 3";
                _cells.GetCell(36, TitleRow03).Value = "SC QTY 3";
                _cells.GetCell(37, TitleRow03).Value = "STOCK CODE 4";
                _cells.GetCell(38, TitleRow03).Value = "SC QTY 4";
                _cells.GetCell(39, TitleRow03).Value = "STOCK CODE 5";
                _cells.GetCell(40, TitleRow03).Value = "SC QTY 5";
                _cells.GetCell(41, TitleRow03).Value = "CANT H.HOMBRE";
                _cells.GetCell(42, TitleRow03).Value = "DURACIÓN TAREA";
                _cells.GetCell(43, TitleRow03).Value = "EQUIPO DETENIDO";
                _cells.GetCell(44, TitleRow03).Value = "RAISED REP.";
                _cells.GetCell(45, TitleRow03).Value = "CAMBIO HORA";
                _cells.GetCell(46, TitleRow03).AddComment("HH:MM");


              
                //
                _cells.GetRange(28, TitleRow01, ResultColumnM01 - 1, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ResultColumnM01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnM01, TitleRow01).Style = StyleConstants.TitleResult;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumnM01, TitleRow01 + 1), TableNameM01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                ////CONSTRUYO LA HOJA 2 - CLOSE WR
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameM02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "CLOSE WORK REQUEST - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = StyleConstants.TitleAdditional;

                _cells.GetRange(1, TitleRow02, ResultColumnM02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow02).Value = "REQUEST ID";
                _cells.GetCell(2, TitleRow02).Value = "CLOSED BY";
                _cells.GetCell(3, TitleRow02).Value = "CLOSED DATE";
                _cells.GetCell(4, TitleRow02).Value = "CLOSED TIME";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ResultColumnM02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumnM02, TitleRow02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow02 + 1, ResultColumnM02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumnM02, TitleRow02 + 1), TableNameM02);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                ////CONSTRUYO LA HOJA 3 RERFERENCE CODES WR
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameM03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "REFERENCE CODES WORK REQUEST - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;
                _cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("K5").Style = StyleConstants.TitleAdditional;

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.GetCell("A5").Value = "STATUS";
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("D4").AddComment("YYYYMMDD");

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell("B4"), ValidationSheetName, 2, false);
                _cells.SetValidationList(_cells.GetCell("B5"), ValidationSheetName, 3, false);
                //Valores predeterminados
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

                _cells.GetRange(1, TitleRow03, ResultColumnM03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;
                for (var i = 4; i < ResultColumnM03; i++)
                {
                    _cells.GetCell(i, TitleRow03 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow03 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow03 - 1).Value = "true";
                }
                _cells.GetRange(1, TitleRow03, ResultColumnM03, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(02, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(03, TitleRow03).Style = StyleConstants.TitleInformation;

                _cells.GetCell(01, TitleRow03).Value = "WORKGROUP";
                _cells.GetCell(02, TitleRow03).Value = "REQUEST ID";
                _cells.GetCell(03, TitleRow03).Value = "DESCRIPTION";

                _cells.GetCell(04, TitleRow01).Value = "SL_AGREEMENT";
                _cells.GetCell(05, TitleRow01).Value = "SLA_FAILURE_CODE";
                _cells.GetCell(06, TitleRow01).Value = "SLA_START_DATE";
                _cells.GetCell(07, TitleRow01).Value = "SLA_START_TIME";
                _cells.GetCell(08, TitleRow01).Value = "SLA_DUE_DATE";
                _cells.GetCell(09, TitleRow01).Value = "SLA_DUE_TIME";
                _cells.GetCell(10, TitleRow01).Value = "SLA_DUE_DAYS";
                _cells.GetCell(11, TitleRow01).Value = "SLA_WARN_DATE";
                _cells.GetCell(12, TitleRow01).Value = "SLA_WARN_TIME";
                _cells.GetCell(13, TitleRow01).Value = "SLA_WARN_DAYS";

                _cells.GetCell(ResultColumnM03, TitleRow03).Value = "RESULTADO";
                _cells.GetCell(ResultColumnM03, TitleRow03).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow03 + 1, ResultColumnM03, TitleRow03 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumnM03, TitleRow03 + 1), TableNameM03);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void ReviewWorkRequestList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRange(TableName01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var workGroup = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var wrStatus = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);

            if (_eFunctions.DebugQueries)
                _cells.GetCell("L1").Value = WorkRequestActions.Queries.GetFetchWorkRequest(_eFunctions.dbReference, _eFunctions.dbLink, workGroup, startDate, endDate, wrStatus);
            var listwr = WorkRequestActions.FetchWorkRequest(_eFunctions, workGroup, startDate, endDate, wrStatus);
            var i = TitleRow01 + 1;
            foreach (var wr in listwr)
            {
                try
                {
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + WrStatusList.GetStatusName(wr.status);
                    _cells.GetCell(04, i).Value = "'" + wr.requestIdDescription1;
                    _cells.GetCell(05, i).Value = "'" + wr.requestIdDescription2;
                    _cells.GetCell(06, i).Value = "'" + wr.equipmentNo;
                                                  
                    //WORK                        
                    _cells.GetCell(07, i).Value = "'" + wr.employee;
                    _cells.GetCell(08, i).Value = "'" + wr.classification;
                    _cells.GetCell(09, i).Value = "'" + wr.requestType;
                    _cells.GetCell(10, i).Value = "'" + wr.userStatus;
                    _cells.GetCell(11, i).Value = "'" + wr.priorityCode;
                    _cells.GetCell(12, i).Value = "'" + wr.region;
                    //SOURCE                      
                    _cells.GetCell(13, i).Value = "'" + wr.contactId;
                    _cells.GetCell(14, i).Value = "'" + wr.source;
                    _cells.GetCell(15, i).Value = "'" + wr.sourceReference;
                    //DATES                       
                    _cells.GetCell(16, i).Value = "'" + wr.requiredByDate;
                    _cells.GetCell(17, i).Value = "'" + wr.requiredByTime;
                    _cells.GetCell(18, i).Value = "'" + wr.raisedUser;
                    _cells.GetCell(19, i).Value = "'" + wr.raisedDate;
                    _cells.GetCell(20, i).Value = "'" + wr.raisedTime;
                    _cells.GetCell(21, i).Value = "'" + wr.closedBy;
                    _cells.GetCell(22, i).Value = "'" + wr.closedDate;
                    //ASSIGN                      
                    _cells.GetCell(23, i).Value = "'" + wr.assignPerson;
                    _cells.GetCell(25, i).Value = "'" + wr.ownerId;
                    //ESTIMATE                    
                    _cells.GetCell(25, i).Value = "'" + wr.estimateNo;
                    _cells.GetCell(26, i).Value = "'" + wr.standardJob;
                    _cells.GetCell(27, i).Value = "'" + wr.standardJobDistrict;
                    //SERVICE LEVEL AGREEMENT     
                    _cells.GetCell(28, i).Value = "'" + wr.ServiceLevelAgreement.ServiceLevel;
                    _cells.GetCell(29, i).Value = "'" + wr.ServiceLevelAgreement.FailureCode;
                    _cells.GetCell(30, i).Value = "'" + wr.ServiceLevelAgreement.StartDate;
                    _cells.GetCell(31, i).Value = "'" + wr.ServiceLevelAgreement.StartTime;
                    _cells.GetCell(32, i).Value = "'" + wr.ServiceLevelAgreement.DueDate;
                    _cells.GetCell(33, i).Value = "'" + wr.ServiceLevelAgreement.DueTime;
                    _cells.GetCell(34, i).Value = "'" + wr.ServiceLevelAgreement.DueDays;
                    _cells.GetCell(35, i).Value = "'" + wr.ServiceLevelAgreement.WarnDate;
                    _cells.GetCell(36, i).Value = "'" + wr.ServiceLevelAgreement.WarnTime;
                    _cells.GetCell(37, i).Value = "'" + wr.ServiceLevelAgreement.WarnDays;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReviewWorkRequestMnttoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableNameM01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var workGroup = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var wrStatus = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);

            if (_eFunctions.DebugQueries)
                _cells.GetCell("L1").Value = WorkRequestActions.Queries.GetFetchWorkRequest(_eFunctions.dbReference, _eFunctions.dbLink, workGroup, startDate, endDate, wrStatus);
            var listwr = WorkRequestActions.FetchWorkRequest(_eFunctions, workGroup, startDate, endDate, wrStatus);
            var i = TitleRow01 + 1;
            foreach (var wr in listwr)
            {
                try
                {
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + WrStatusList.GetStatusName(wr.status);
                    _cells.GetCell(04, i).Value = "'" + wr.requestIdDescription1;
                    _cells.GetCell(05, i).Value = "'" + wr.requestIdDescription2;
                    _cells.GetCell(06, i).Value = "'" + wr.equipmentNo;

                    //WORK                        
                    _cells.GetCell(07, i).Value = "'" + wr.employee;
                    _cells.GetCell(08, i).Value = "'" + wr.classification;
                    _cells.GetCell(09, i).Value = "'" + wr.requestType;
                    _cells.GetCell(10, i).Value = "'" + wr.userStatus;
                    _cells.GetCell(11, i).Value = "'" + wr.priorityCode;
                    _cells.GetCell(12, i).Value = "'" + wr.region;
                    //SOURCE                      
                    _cells.GetCell(13, i).Value = "'" + wr.contactId;
                    _cells.GetCell(14, i).Value = "'" + wr.source;
                    _cells.GetCell(15, i).Value = "'" + wr.sourceReference;
                    //DATES                       
                    _cells.GetCell(16, i).Value = "'" + wr.requiredByDate;
                    _cells.GetCell(17, i).Value = "'" + wr.requiredByTime;
                    _cells.GetCell(18, i).Value = "'" + wr.raisedUser;
                    _cells.GetCell(19, i).Value = "'" + wr.raisedDate;
                    _cells.GetCell(20, i).Value = "'" + wr.raisedTime;
                    _cells.GetCell(21, i).Value = "'" + wr.closedBy;
                    _cells.GetCell(22, i).Value = "'" + wr.closedDate;
                    //ASSIGN                      
                    _cells.GetCell(23, i).Value = "'" + wr.assignPerson;
                    _cells.GetCell(25, i).Value = "'" + wr.ownerId;
                    //ESTIMATE                    
                    _cells.GetCell(25, i).Value = "'" + wr.estimateNo;
                    _cells.GetCell(26, i).Value = "'" + wr.standardJob;
                    _cells.GetCell(27, i).Value = "'" + wr.standardJobDistrict;
                    //REFERENCE CODES    
                    _cells.GetCell(28, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionHeader;
                    _cells.GetCell(29, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionBody;
                    _cells.GetCell(30, i).Value = "'" + wr.ReferenceCodes.WorkOrderOrigen;
                    _cells.GetCell(31, i).Value = "'" + wr.ReferenceCodes.StockCode1;
                    _cells.GetCell(32, i).Value = "'" + wr.ReferenceCodes.StockQuantity1;
                    _cells.GetCell(33, i).Value = "'" + wr.ReferenceCodes.StockCode2;
                    _cells.GetCell(34, i).Value = "'" + wr.ReferenceCodes.StockQuantity2;
                    _cells.GetCell(35, i).Value = "'" + wr.ReferenceCodes.StockCode3;
                    _cells.GetCell(36, i).Value = "'" + wr.ReferenceCodes.StockQuantity3;
                    _cells.GetCell(37, i).Value = "'" + wr.ReferenceCodes.StockCode4;
                    _cells.GetCell(38, i).Value = "'" + wr.ReferenceCodes.StockQuantity4;
                    _cells.GetCell(39, i).Value = "'" + wr.ReferenceCodes.StockCode5;
                    _cells.GetCell(40, i).Value = "'" + wr.ReferenceCodes.StockQuantity5;
                    _cells.GetCell(41, i).Value = "'" + wr.ReferenceCodes.HorasHombre;
                    _cells.GetCell(42, i).Value = "'" + wr.ReferenceCodes.DuracionTarea;
                    _cells.GetCell(43, i).Value = "'" + wr.ReferenceCodes.EquipoDetenido;
                    _cells.GetCell(44, i).Value = "'" + wr.ReferenceCodes.RaisedReprogramada;
                    _cells.GetCell(45, i).Value = "'" + wr.ReferenceCodes.CambioHora;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReReviewWorkRequestList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);
                    if (_eFunctions.DebugQueries)
                        _cells.GetCell("L1").Value = WorkRequestActions.Queries.GetFetchWorkRequest(_eFunctions.dbReference, _eFunctions.dbLink, requestId);
                    if (wr == null || wr.requestId == null)
                        throw new Exception("WORK REQUEST NO ENCONTRADO");
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + WrStatusList.GetStatusName(wr.status);
                    _cells.GetCell(04, i).Value = "'" + wr.requestIdDescription1;
                    _cells.GetCell(05, i).Value = "'" + wr.requestIdDescription2;
                    _cells.GetCell(06, i).Value = "'" + wr.equipmentNo;

                    //WORK                        
                    _cells.GetCell(07, i).Value = "'" + wr.employee;
                    _cells.GetCell(08, i).Value = "'" + wr.classification;
                    _cells.GetCell(09, i).Value = "'" + wr.requestType;
                    _cells.GetCell(10, i).Value = "'" + wr.userStatus;
                    _cells.GetCell(11, i).Value = "'" + wr.priorityCode;
                    _cells.GetCell(12, i).Value = "'" + wr.region;
                    //SOURCE                      
                    _cells.GetCell(13, i).Value = "'" + wr.contactId;
                    _cells.GetCell(14, i).Value = "'" + wr.source;
                    _cells.GetCell(15, i).Value = "'" + wr.sourceReference;
                    //DATES                       
                    _cells.GetCell(16, i).Value = "'" + wr.requiredByDate;
                    _cells.GetCell(17, i).Value = "'" + wr.requiredByTime;
                    _cells.GetCell(18, i).Value = "'" + wr.raisedUser;
                    _cells.GetCell(19, i).Value = "'" + wr.raisedDate;
                    _cells.GetCell(20, i).Value = "'" + wr.raisedTime;
                    _cells.GetCell(21, i).Value = "'" + wr.closedBy;
                    _cells.GetCell(22, i).Value = "'" + wr.closedDate;
                    //ASSIGN                      
                    _cells.GetCell(23, i).Value = "'" + wr.assignPerson;
                    _cells.GetCell(25, i).Value = "'" + wr.ownerId;
                    //ESTIMATE                    
                    _cells.GetCell(25, i).Value = "'" + wr.estimateNo;
                    _cells.GetCell(26, i).Value = "'" + wr.standardJob;
                    _cells.GetCell(27, i).Value = "'" + wr.standardJobDistrict;
                    //SERVICE LEVEL AGREEMENT     
                    _cells.GetCell(28, i).Value = "'" + wr.ServiceLevelAgreement.ServiceLevel;
                    _cells.GetCell(29, i).Value = "'" + wr.ServiceLevelAgreement.FailureCode;
                    _cells.GetCell(30, i).Value = "'" + wr.ServiceLevelAgreement.StartDate;
                    _cells.GetCell(31, i).Value = "'" + wr.ServiceLevelAgreement.StartTime;
                    _cells.GetCell(32, i).Value = "'" + wr.ServiceLevelAgreement.DueDate;
                    _cells.GetCell(33, i).Value = "'" + wr.ServiceLevelAgreement.DueTime;
                    _cells.GetCell(34, i).Value = "'" + wr.ServiceLevelAgreement.DueDays;
                    _cells.GetCell(35, i).Value = "'" + wr.ServiceLevelAgreement.WarnDate;
                    _cells.GetCell(36, i).Value = "'" + wr.ServiceLevelAgreement.WarnTime;
                    _cells.GetCell(37, i).Value = "'" + wr.ServiceLevelAgreement.WarnDays;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReReviewWorkRequestMnttoList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameM01, ResultColumnM01);

            var i = TitleRow01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);
                    if (_eFunctions.DebugQueries)
                        _cells.GetCell("L1").Value = WorkRequestActions.Queries.GetFetchWorkRequest(_eFunctions.dbReference, _eFunctions.dbLink, requestId);
                    if (wr == null || wr.requestId == null)
                        throw new Exception("WORK REQUEST NO ENCONTRADO");
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + WrStatusList.GetStatusName(wr.status);
                    _cells.GetCell(04, i).Value = "'" + wr.requestIdDescription1;
                    _cells.GetCell(05, i).Value = "'" + wr.requestIdDescription2;
                    _cells.GetCell(06, i).Value = "'" + wr.equipmentNo;

                    //WORK                        
                    _cells.GetCell(07, i).Value = "'" + wr.employee;
                    _cells.GetCell(08, i).Value = "'" + wr.classification;
                    _cells.GetCell(09, i).Value = "'" + wr.requestType;
                    _cells.GetCell(10, i).Value = "'" + wr.userStatus;
                    _cells.GetCell(11, i).Value = "'" + wr.priorityCode;
                    _cells.GetCell(12, i).Value = "'" + wr.region;
                    //SOURCE                      
                    _cells.GetCell(13, i).Value = "'" + wr.contactId;
                    _cells.GetCell(14, i).Value = "'" + wr.source;
                    _cells.GetCell(15, i).Value = "'" + wr.sourceReference;
                    //DATES                       
                    _cells.GetCell(16, i).Value = "'" + wr.requiredByDate;
                    _cells.GetCell(17, i).Value = "'" + wr.requiredByTime;
                    _cells.GetCell(18, i).Value = "'" + wr.raisedUser;
                    _cells.GetCell(19, i).Value = "'" + wr.raisedDate;
                    _cells.GetCell(20, i).Value = "'" + wr.raisedTime;
                    _cells.GetCell(21, i).Value = "'" + wr.closedBy;
                    _cells.GetCell(22, i).Value = "'" + wr.closedDate;
                    //ASSIGN                      
                    _cells.GetCell(23, i).Value = "'" + wr.assignPerson;
                    _cells.GetCell(25, i).Value = "'" + wr.ownerId;
                    //ESTIMATE                    
                    _cells.GetCell(25, i).Value = "'" + wr.estimateNo;
                    _cells.GetCell(26, i).Value = "'" + wr.standardJob;
                    _cells.GetCell(27, i).Value = "'" + wr.standardJobDistrict;
                    //REFERENCE CODES     
                    _cells.GetCell(28, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionHeader;
                    _cells.GetCell(29, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionBody;
                    _cells.GetCell(30, i).Value = "'" + wr.ReferenceCodes.WorkOrderOrigen;
                    _cells.GetCell(31, i).Value = "'" + wr.ReferenceCodes.StockCode1;
                    _cells.GetCell(32, i).Value = "'" + wr.ReferenceCodes.StockQuantity1;
                    _cells.GetCell(33, i).Value = "'" + wr.ReferenceCodes.StockCode2;
                    _cells.GetCell(34, i).Value = "'" + wr.ReferenceCodes.StockQuantity2;
                    _cells.GetCell(35, i).Value = "'" + wr.ReferenceCodes.StockCode3;
                    _cells.GetCell(36, i).Value = "'" + wr.ReferenceCodes.StockQuantity3;
                    _cells.GetCell(37, i).Value = "'" + wr.ReferenceCodes.StockCode4;
                    _cells.GetCell(38, i).Value = "'" + wr.ReferenceCodes.StockQuantity4;
                    _cells.GetCell(39, i).Value = "'" + wr.ReferenceCodes.StockCode5;
                    _cells.GetCell(40, i).Value = "'" + wr.ReferenceCodes.StockQuantity5;
                    _cells.GetCell(41, i).Value = "'" + wr.ReferenceCodes.HorasHombre;
                    _cells.GetCell(42, i).Value = "'" + wr.ReferenceCodes.DuracionTarea;
                    _cells.GetCell(43, i).Value = "'" + wr.ReferenceCodes.EquipoDetenido;
                    _cells.GetCell(44, i).Value = "'" + wr.ReferenceCodes.RaisedReprogramada;
                    _cells.GetCell(45, i).Value = "'" + wr.ReferenceCodes.CambioHora;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }
        private void ReviewReferenceCodesList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRange(TableName03);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var workGroup = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var wrStatus = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);

            if (_eFunctions.DebugQueries)
                _cells.GetCell("L1").Value = WorkRequestActions.Queries.GetFetchWorkRequest(_eFunctions.dbReference, _eFunctions.dbLink, workGroup, startDate, endDate, wrStatus);
            var listwr = WorkRequestActions.FetchWorkRequest(_eFunctions, workGroup, startDate, endDate, wrStatus);
            var i = TitleRow03 + 1;
            foreach (var wr in listwr)
            {
                try
                {
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + wr.requestIdDescription1 + " " + wr.requestIdDescription2;
                    _cells.GetCell(04, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionHeader;
                    _cells.GetCell(05, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionBody;
                    _cells.GetCell(06, i).Value = "'" + wr.ReferenceCodes.WorkOrderOrigen;
                    _cells.GetCell(07, i).Value = "'" + wr.ReferenceCodes.StockCode1;
                    _cells.GetCell(08, i).Value = "'" + wr.ReferenceCodes.StockQuantity1;
                    _cells.GetCell(09, i).Value = "'" + wr.ReferenceCodes.StockCode2;
                    _cells.GetCell(10, i).Value = "'" + wr.ReferenceCodes.StockQuantity2;
                    _cells.GetCell(11, i).Value = "'" + wr.ReferenceCodes.StockCode3;
                    _cells.GetCell(12, i).Value = "'" + wr.ReferenceCodes.StockQuantity3;
                    _cells.GetCell(13, i).Value = "'" + wr.ReferenceCodes.StockCode4;
                    _cells.GetCell(14, i).Value = "'" + wr.ReferenceCodes.StockQuantity4;
                    _cells.GetCell(15, i).Value = "'" + wr.ReferenceCodes.StockCode5;
                    _cells.GetCell(16, i).Value = "'" + wr.ReferenceCodes.StockQuantity5;
                    _cells.GetCell(17, i).Value = "'" + wr.ReferenceCodes.HorasHombre;
                    _cells.GetCell(18, i).Value = "'" + wr.ReferenceCodes.HorasQty;
                    _cells.GetCell(19, i).Value = "'" + wr.ReferenceCodes.DuracionTarea;
                    _cells.GetCell(20, i).Value = "'" + wr.ReferenceCodes.EquipoDetenido;
                    _cells.GetCell(21, i).Value = "'" + wr.ReferenceCodes.RaisedReprogramada;
                    _cells.GetCell(22, i).Value = "'" + wr.ReferenceCodes.CambioHora;

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewReferenceCodesList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            //porque son columnas extensas
            _cells.GetCell(04, 01).ColumnWidth = 30;
            _cells.GetCell(05, 01).ColumnWidth = 30;
            _cells?.SetCursorDefault();
        }
        private void ReReviewReferenceCodesList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName03, ResultColumn03);

            var i = TitleRow03 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);
                    if (_eFunctions.DebugQueries)
                        _cells.GetCell("L1").Value = WorkRequestActions.Queries.GetFetchWorkRequest(_eFunctions.dbReference, _eFunctions.dbLink, requestId);

                    if(wr == null || wr.requestId == null)
                        throw new Exception ("WORK REQUEST NO ENCONTRADO");
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + wr.requestIdDescription1 + " " + wr.requestIdDescription2;
                    _cells.GetCell(04, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionHeader;
                    _cells.GetCell(05, i).Value = "'" + wr.ReferenceCodes.ExtendedDescriptionBody;
                    _cells.GetCell(06, i).Value = "'" + wr.ReferenceCodes.WorkOrderOrigen;
                    _cells.GetCell(07, i).Value = "'" + wr.ReferenceCodes.StockCode1;
                    _cells.GetCell(08, i).Value = "'" + wr.ReferenceCodes.StockQuantity1;
                    _cells.GetCell(09, i).Value = "'" + wr.ReferenceCodes.StockCode2;
                    _cells.GetCell(10, i).Value = "'" + wr.ReferenceCodes.StockQuantity2;
                    _cells.GetCell(11, i).Value = "'" + wr.ReferenceCodes.StockCode3;
                    _cells.GetCell(12, i).Value = "'" + wr.ReferenceCodes.StockQuantity3;
                    _cells.GetCell(13, i).Value = "'" + wr.ReferenceCodes.StockCode4;
                    _cells.GetCell(14, i).Value = "'" + wr.ReferenceCodes.StockQuantity4;
                    _cells.GetCell(15, i).Value = "'" + wr.ReferenceCodes.StockCode5;
                    _cells.GetCell(16, i).Value = "'" + wr.ReferenceCodes.StockQuantity5;
                    _cells.GetCell(17, i).Value = "'" + wr.ReferenceCodes.HorasHombre;
                    _cells.GetCell(18, i).Value = "'" + wr.ReferenceCodes.HorasQty;
                    _cells.GetCell(19, i).Value = "'" + wr.ReferenceCodes.DuracionTarea;
                    _cells.GetCell(20, i).Value = "'" + wr.ReferenceCodes.EquipoDetenido;
                    _cells.GetCell(21, i).Value = "'" + wr.ReferenceCodes.RaisedReprogramada;
                    _cells.GetCell(22, i).Value = "'" + wr.ReferenceCodes.CambioHora;

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewReferenceCodesList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
            //porque son columnas extensas
            _cells.GetCell(04, 01).ColumnWidth = 30;
            _cells.GetCell(05, 01).ColumnWidth = 30;
        }
        private void CreateWorkRequestList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        requestId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value),

                        requestIdDescription1 = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value),
                        requestIdDescription2 = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value),
                        equipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),

                        employee = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        classification = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        requestType = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value),
                        userStatus = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                        priorityCode = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value),
                        region = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value),

                        contactId = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value),
                        source = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value),
                        sourceReference = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value),

                        requiredByDate = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value),
                        requiredByTime = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value),
                        raisedUser = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value),
                        raisedDate = _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value),
                        raisedTime = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value),
                        closedBy = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value),
                        closedDate = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value),

                        assignPerson = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value),
                        ownerId = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value),

                        estimateNo = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value),
                        standardJob = _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value),
                        standardJobDistrict = _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value),

                        ServiceLevelAgreement =
                        {
                            ServiceLevel = _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value),
                            FailureCode = _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value),
                            StartDate = _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value),
                            StartTime = _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value),
                            DueDate = _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value),
                            DueTime = _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value),
                            DueDays = _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value),
                            WarnDate = _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value),
                            WarnTime = _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value),
                            WarnDays = _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value)
                        }
                    };

                    var replySheet = WorkRequestActions.CreateWorkRequest(urlService, opSheet, wr);
                    var requestId = replySheet.requestId;
                    _cells.GetCell(2, i).Value = "'" + requestId;
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    if (!string.IsNullOrWhiteSpace("" + wr.ServiceLevelAgreement.ServiceLevel))
                        WorkRequestActions.SetWorkRequestSla(urlService, opSheet, requestId, wr.ServiceLevelAgreement);
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Value = "CREADO " + requestId;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void CreateWorkRequestMnttoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameM01, ResultColumnM01);

            var i = TitleRow01 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        requestId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value),

                        requestIdDescription1 = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value),
                        requestIdDescription2 = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value),
                        equipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),

                        employee = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        classification = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        requestType = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value),
                        userStatus = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                        priorityCode = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value),
                        region = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value),

                        contactId = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value),
                        source = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value),
                        sourceReference = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value),

                        requiredByDate = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value),
                        requiredByTime = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value),
                        raisedUser = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value),
                        raisedDate = _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value),
                        raisedTime = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value),
                        closedBy = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value),
                        closedDate = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value),

                        assignPerson = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value),
                        ownerId = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value),

                        estimateNo = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value),
                        standardJob = _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value),
                        standardJobDistrict = _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value),
                        
                        ReferenceCodes =
                        {
                            ExtendedDescriptionHeader=  _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value),
                            ExtendedDescriptionBody=  _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value),
                            WorkOrderOrigen=  _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value),
                            StockCode1 =  _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value),
                            StockQuantity1= _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value),
                            StockCode2= _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value),
                            StockQuantity2 =  _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value),
                            StockCode3= _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value),
                            StockQuantity3=  _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value),
                            StockCode4= _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value),
                            StockQuantity4=  _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value),
                            StockCode5  = _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value),
                            StockQuantity5=  _cells.GetEmptyIfNull(_cells.GetCell(40, i).Value),
                            HorasHombre=  _cells.GetEmptyIfNull(_cells.GetCell(41, i).Value),
                            DuracionTarea=  _cells.GetEmptyIfNull(_cells.GetCell(42, i).Value),
                            EquipoDetenido=  _cells.GetEmptyIfNull(_cells.GetCell(43, i).Value),
                            RaisedReprogramada=  _cells.GetEmptyIfNull(_cells.GetCell(44, i).Value),
                            CambioHora=  _cells.GetEmptyIfNull(_cells.GetCell(45, i).Value)
                        }
                    };

                    var replySheet = WorkRequestActions.CreateWorkRequest(urlService, opSheet, wr);
                    var requestId = replySheet.requestId;
                    if (string.IsNullOrWhiteSpace(replySheet.requestId))
                        throw new Exception("No se ha podido crear el WorkRequest");
                    //Creacion de los reference codes
                    var refCodeOpContext = ReferenceCodeActions.GetRefCodesOpContext(opSheet.district, opSheet.position, opSheet.maxInstances, opSheet.returnWarnings, opSheet.returnWarningsSpecified);
                    var stdTextOpContext = StdText.GetCustomOpContext(opSheet.district, opSheet.position, opSheet.maxInstances, opSheet.returnWarnings);
                    var error = "";
                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.StockCode1))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "001", "9001", wr.ReferenceCodes.StockCode1);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (!string.IsNullOrWhiteSpace(stdTextId))
                            StdText.SetCustomText(urlService, stdTextOpContext, stdTextId, wr.ReferenceCodes.StockQuantity1);
                        else
                            error += " / Error al crear SC1";
                    }
                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.StockCode2))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "001", "9002", wr.ReferenceCodes.StockCode2);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (!string.IsNullOrWhiteSpace(stdTextId))
                            StdText.SetCustomText(urlService, stdTextOpContext, stdTextId, wr.ReferenceCodes.StockQuantity2);
                        else
                            error += " / Error al crear SC2";
                    }
                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.StockCode3))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "001", "9003", wr.ReferenceCodes.StockCode3);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (!string.IsNullOrWhiteSpace(stdTextId))
                            StdText.SetCustomText(urlService, stdTextOpContext, stdTextId, wr.ReferenceCodes.StockQuantity3);
                        else
                            error += " / Error al crear SC3";
                    }
                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.StockCode4))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "001", "9004", wr.ReferenceCodes.StockCode4);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (!string.IsNullOrWhiteSpace(stdTextId))
                            StdText.SetCustomText(urlService, stdTextOpContext, stdTextId, wr.ReferenceCodes.StockQuantity4);
                        else
                            error += " / Error al crear SC4";
                    }
                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.StockCode5))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "001", "9005", wr.ReferenceCodes.StockCode5);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (!string.IsNullOrWhiteSpace(stdTextId))
                            StdText.SetCustomText(urlService, stdTextOpContext, stdTextId, wr.ReferenceCodes.StockQuantity5);
                        else
                            error += " / Error al crear SC5";
                    }

                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.HorasHombre))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "006", "001", wr.ReferenceCodes.HorasHombre);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (!string.IsNullOrWhiteSpace(stdTextId))
                            StdText.SetCustomText(urlService, stdTextOpContext, stdTextId, wr.ReferenceCodes.HorasQty);
                        else
                            error += " / Error al crear HH";
                    }

                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.DuracionTarea))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "007", "001", wr.ReferenceCodes.DuracionTarea);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (string.IsNullOrWhiteSpace(stdTextId))
                            error += " / Error al crear Duracion";
                    }

                    if (!string.IsNullOrWhiteSpace(wr.ReferenceCodes.WorkOrderOrigen))
                    {
                        var refItem = new ReferenceCodeItem("WRQ", requestId, "009", "001", wr.ReferenceCodes.WorkOrderOrigen);
                        var replyRefCode = ReferenceCodeActions.ModifyRefCode(urlService, refCodeOpContext, refItem);
                        var stdTextId = replyRefCode.stdTxtKey;
                        if (string.IsNullOrWhiteSpace(stdTextId))
                            error += " / Error al crear OT Origen";
                    }
                    //
                    _cells.GetCell(2, i).Value = "'" + requestId;
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    if (!string.IsNullOrWhiteSpace("" + wr.ServiceLevelAgreement.ServiceLevel))
                        WorkRequestActions.SetWorkRequestSla(urlService, opSheet, requestId, wr.ServiceLevelAgreement);
                    _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnM01, i).Value = "CREADO " + requestId;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ModifyWorkRequestList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup = Utils.IsTrue(_cells.GetCell(1, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value) : null,
                        requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        requestIdDescription1 = Utils.IsTrue(_cells.GetCell(4, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value) : null,
                        requestIdDescription2 = Utils.IsTrue(_cells.GetCell(5, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value) : null,
                        equipmentNo = Utils.IsTrue(_cells.GetCell(6, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value) : null,
                        employee = Utils.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null,
                        classification = Utils.IsTrue(_cells.GetCell(8, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value) : null,
                        requestType = Utils.IsTrue(_cells.GetCell(9, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value) : null,
                        userStatus = Utils.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null,
                        priorityCode = Utils.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null,
                        region = Utils.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null,
                        contactId = Utils.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null,
                        source = Utils.IsTrue(_cells.GetCell(14, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null,
                        sourceReference = Utils.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null,
                        requiredByDate = Utils.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null,
                        requiredByTime = Utils.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null,
                        raisedUser = Utils.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null,
                        raisedDate = Utils.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null,
                        raisedTime = Utils.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null,
                        closedBy = Utils.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null,
                        closedDate = Utils.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null,
                        assignPerson = Utils.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null,
                        ownerId = Utils.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null,
                        estimateNo = Utils.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null,
                        standardJob = Utils.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null,
                        standardJobDistrict = Utils.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null,
                        ServiceLevelAgreement =
                        {
                            ServiceLevel = Utils.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null,
                            FailureCode = Utils.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null,
                            StartDate = Utils.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null,
                            StartTime = Utils.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null,
                            DueDate = Utils.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null,
                            DueTime = Utils.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null,
                            DueDays = Utils.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null,
                            WarnDate = Utils.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null,
                            WarnTime = Utils.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null,
                            WarnDays = Utils.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null
                        }
                    };

                    WorkRequestActions.ModifyWorkRequest(urlService, opSheet, wr);
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void CloseWorkRequestList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);

            var i = TitleRow02 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    var closedBy = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var closedDate = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var closedTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);

                    WorkRequestActions.CloseWorkRequest(urlService, opSheet, requestId, closedBy, closedDate, closedTime);
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Value = "CERRADO";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CloseWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReOpenWorkRequestList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);

            var i = TitleRow02 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);

                    WorkRequestActions.ReOpenWorkRequest(urlService, opSheet, requestId);
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Value = "REABIERTA";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReOpenWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }
        private void DeleteWorkRequestList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);

                    WorkRequestActions.DeleteWorkRequest(urlService, opSheet, requestId);
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Value = "ELIMINADO";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:DeleteWorkRequestList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void SetSlaList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var serviceLevelAgreement = new ServiceLevelAgreement
                    {
                            ServiceLevel = Utils.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null,
                            FailureCode = Utils.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null,
                            StartDate = Utils.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null,
                            StartTime = Utils.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null,
                            DueDate = Utils.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null,
                            DueTime = Utils.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null,
                            DueDays = Utils.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null,
                            WarnDate = Utils.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null,
                            WarnTime = Utils.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null,
                            WarnDays = Utils.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null
                    };
                    
                    WorkRequestActions.SetWorkRequestSla(urlService, opSheet, requestId, serviceLevelAgreement);
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Value = "SLA ESTABLECIDO";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:SetSlaList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ResetSlaList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = _eFunctions.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var serviceLevelAgreement = new ServiceLevelAgreement
                    {
                        ServiceLevel = Utils.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null,
                        FailureCode = Utils.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null,
                        StartDate = Utils.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null,
                        StartTime = Utils.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null,
                        DueDate = Utils.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null,
                        DueTime = Utils.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null,
                        DueDays = Utils.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null,
                        WarnDate = Utils.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null,
                        WarnTime = Utils.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null,
                        WarnDays = Utils.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null
                    };

                    WorkRequestActions.ResetWorkRequestSla(urlService, opSheet, requestId, serviceLevelAgreement);
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Value = "SLA RESETEADO";

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ResetSlaList()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }

     
    }
}
