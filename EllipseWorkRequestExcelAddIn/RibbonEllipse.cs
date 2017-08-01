using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseWorkRequestClassLibrary;
using EllipseWorkRequestClassLibrary.WorkRequestService;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseWorkRequestExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "WorkRequest";
        private const string SheetName02 = "WorkRequestClose";
        private const string SheetName03 = "WorkRequestsReferences";
        private const string SheetNameM01 = "WorkRequestMntto";
        private const string SheetNameM02 = "WorkRequestMnttoClose";
        private const string SheetNameM03 = "WorkRequestsMnttoSLA";
        private const string SheetNameV01 = "WorkRequestVagones";
        private const string SheetNamePfc01 = "WorkRequestSolicitudesFC";
        //private const string SheetName04 = "WorkOrdersRelated";
        private const int TitleRow01 = 9;
        private const int TitleRow02 = 6;
        private const int TitleRow03 = 9;
        private const int TitleRowM02 = 6;
        private const int TitleRowM01 = 9;
        private const int TitleRowV01 = 5;
        private const int TitleRowPfc01 = 5;
        private const int ResultColumn01 = 38;
        private const int ResultColumn02 = 5;
        private const int ResultColumn03 = 23;
        private const int ResultColumnM01 = 46;
        private const int ResultColumnM02 = 5;
        private const int ResultColumnV01 = 11;
        private const int ResultColumnPfc01 = 10;
        //private const int ResultColumnM03 = 14;
        private const string TableName01 = "WorkRequestTable";
        private const string TableName02 = "WorkRequestCloseTable";
        private const string TableName03 = "WorkRequestsReferencesTable";
        private const string TableNameM01 = "WorkRequestTable";
        private const string TableNameM02 = "WorkRequestCloseTable";
        private const string TableNameM03 = "WorkRequestSLATable";
        private const string TableNameV01 = "WorkRequestVagonesTable";
        private const string TableNamePfc01 = "WorkRequestSolicitudesFCTable";
        //private const string TableName04 = "WorkOrdersRelatedTable";
        private const string ValidationSheetName = "ValidationSheet";
        private ExcelStyleCells _cells;
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private Application _excelApp;
        // ReSharper disable once FieldCanBeMadeReadOnly.Local
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

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
                Debugger.LogError("RibbonEllipse:ReviewWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
                    _thread = new Thread(ReReviewWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReReviewWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameV01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(CreateWorkRequestVagonesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNamePfc01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(CreateWorkRequestPfcList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void CreateWorkRequestPfcList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNamePfc01, ResultColumnPfc01);

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowPfc01 + 1;
            //default values
            var todayDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) +
                            string.Format("{0:00}", DateTime.Now.Day);
            //To Do change for ICARROS Group Admin
            var employee = _frmAuth.EllipseUser;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup = "PLANFC",
                        requestId = null,
                        requestIdDescription1 = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        requestIdDescription2 = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        equipmentNo = "FERROCARRIL",
                        employee = string.IsNullOrEmpty(_cells.GetEmptyIfNull(_cells.GetCell(4, i).Value))
                                ? employee
                                : _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value),
                        classification = "SS",
                        requestType = "ES",
                        priorityCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(5, i).Value)),
                        contactId = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),
                        sourceReference = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        raisedDate = string.IsNullOrWhiteSpace(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value))
                                    ? todayDate
                                    : _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        ServiceLevelAgreement =
                        {
                            ServiceLevel = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)),
                            StartDate = todayDate
                        }
                    };

                    if (string.IsNullOrWhiteSpace(wr.ServiceLevelAgreement.ServiceLevel) ||
                        string.IsNullOrWhiteSpace(wr.ServiceLevelAgreement.StartDate))
                        throw new Exception("No se puede crear Work Request. Falta la información del Service Level");
                    var replySheet = WorkRequestActions.CreateWorkRequest(urlService, opSheet, wr);
                    var requestId = replySheet.requestId;

                    WorkRequestActions.SetWorkRequestSla(urlService, opSheet, requestId, wr.ServiceLevelAgreement);
                    _cells.GetCell(ResultColumnPfc01, i).Style = StyleConstants.Success;
                    _cells.GetCell(01, i).Value = requestId;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnPfc01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnPfc01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWorkRequestVagonesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumnPfc01, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnModifyWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01))
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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(ModifyWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ModifyWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCloseWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
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
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(CloseWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CloseWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReOpenWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName02))
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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReOpenWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DeleteWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnDeleteWorkRequest_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var dr =
                    MessageBox.Show(
                        @"Esta acción eliminará los WorkRequest existentes en la hoja. ¿Está seguro que desea continuar?",
                        @"ELIMINAR WORK REQUEST", MessageBoxButtons.YesNo);
                if (dr != DialogResult.Yes)
                    return;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01))
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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(DeleteWorkRequestMnttoList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DeleteWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnSetSla_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01))
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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01) ||
                         _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM02) ||
                         _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM03))
                {
                    throw new NotImplementedException(
                        "Las acciones de SLA no están disponibles para el format de Mantenimiento");
                    //_frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    //_frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    //if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    ////si ya hay un thread corriendo que no se ha detenido
                    //if (_thread != null && _thread.IsAlive) return;
                    //_thread = new Thread(SetSlaList);
                    //_thread.SetApartmentState(ApartmentState.STA);
                    //_thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:SetSlaList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnResetSla_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01))
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
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01) ||
                         _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM02) ||
                         _excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM03))
                {
                    throw new NotImplementedException(
                        "Las acciones de SLA no están disponibles para el format de Mantenimiento");
                    //_frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    //_frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    //if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    ////si ya hay un thread corriendo que no se ha detenido
                    //if (_thread != null && _thread.IsAlive) return;
                    //_thread = new Thread(ResetSlaList);
                    //_thread.SetApartmentState(ApartmentState.STA);
                    //_thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:SetSlaList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReviewRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName03))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewReferenceCodesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01))
                    MessageBox.Show(
                        @"Para los Reference Codes de Mantenimiento, utilice las acciones del menú principal Work Request");
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewReferenceCodesList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReReviewRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName03))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewReferenceCodesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01))
                    MessageBox.Show(
                        @"Para los Reference Codes de Mantenimiento, utilice las acciones del menú principal Work Request");
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReReviewReferenceCodesList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName03))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateReferenceCodesList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetNameM01))
                    MessageBox.Show(
                        @"Para los Reference Codes de Mantenimiento, utilice las acciones del menú principal Work Request");
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:UpdateReferenceCodesList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
                _cells.ClearTableRangeColumn(TableName03, ResultColumn03);

                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                var opContext = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var i = TitleRow03 + 1;
                const int validationRow = TitleRow03 - 1;

                while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
                {
                    try
                    {
                        var requestId = _cells.GetEmptyIfNull(_cells.GetCell(02, i).Value);
                        var wr = new WorkRequest();
                        var header = Utils.IsTrue(_cells.GetCell(04, validationRow).Value)
                            ? _cells.GetEmptyIfNull(_cells.GetCell(04, i).Value)
                            : null;
                        var body = Utils.IsTrue(_cells.GetCell(05, validationRow).Value)
                            ? _cells.GetEmptyIfNull(_cells.GetCell(05, i).Value)
                            : null;
                        wr.SetExtendedDescription(header, body);
                        var wrRefCodes = new WorkRequestReferenceCodes
                        {
                            WorkOrderOrigen =
                                Utils.IsTrue(_cells.GetCell(06, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(06, i).Value)
                                    : null,
                            StockCode1 =
                                Utils.IsTrue(_cells.GetCell(07, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(07, i).Value)
                                    : null,
                            StockCode1Qty =
                                Utils.IsTrue(_cells.GetCell(08, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(08, i).Value)
                                    : null,
                            StockCode2 =
                                Utils.IsTrue(_cells.GetCell(09, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(09, i).Value)
                                    : null,
                            StockCode2Qty =
                                Utils.IsTrue(_cells.GetCell(10, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value)
                                    : null,
                            StockCode3 =
                                Utils.IsTrue(_cells.GetCell(11, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value)
                                    : null,
                            StockCode3Qty =
                                Utils.IsTrue(_cells.GetCell(12, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value)
                                    : null,
                            StockCode4 =
                                Utils.IsTrue(_cells.GetCell(13, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value)
                                    : null,
                            StockCode4Qty =
                                Utils.IsTrue(_cells.GetCell(14, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)
                                    : null,
                            StockCode5 =
                                Utils.IsTrue(_cells.GetCell(15, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value)
                                    : null,
                            StockCode5Qty =
                                Utils.IsTrue(_cells.GetCell(16, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value)
                                    : null,
                            HorasHombre =
                                Utils.IsTrue(_cells.GetCell(17, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value)
                                    : null,
                            HorasQty =
                                Utils.IsTrue(_cells.GetCell(18, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value)
                                    : null,
                            DuracionTarea =
                                Utils.IsTrue(_cells.GetCell(19, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value)
                                    : null,
                            EquipoDetenido =
                                Utils.IsTrue(_cells.GetCell(20, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value)
                                    : null,
                            RaisedReprogramada =
                                Utils.IsTrue(_cells.GetCell(21, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value)
                                    : null,
                            CambioHora =
                                Utils.IsTrue(_cells.GetCell(22, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value)
                                    : null
                        };

                        var errorList = "";
                        var replyExtended = WorkRequestActions.UpdateWorkRequestExtendedDescription(urlService,
                            opContext, requestId, wr.GetExtendedDescription(urlService, opContext));
                        if (replyExtended != null && replyExtended.Errors != null && replyExtended.Errors.Length > 0)
                            foreach (var error in replyExtended.Errors)
                                errorList += "\nError: " + error;

                        var replyRefCode = WorkRequestReferenceCodesActions.ModifyReferenceCodes(_eFunctions, urlService,
                            opContext, requestId, wrRefCodes);
                        if (replyRefCode != null && replyRefCode.Errors != null && replyRefCode.Errors.Length > 0)
                            foreach (var error in replyExtended.Errors)
                                errorList += "\nError: " + error;

                        if (!string.IsNullOrWhiteSpace(errorList))
                        {
                            _cells.GetCell(2, i).Value = "'" + requestId;
                            _cells.GetCell(2, i).Style = StyleConstants.Success;

                            _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Warning;
                            _cells.GetCell(ResultColumn03, i).Value = "ACTUALIZADO " + errorList;
                        }
                        else
                        {
                            _cells.GetCell(2, i).Value = "'" + requestId;
                            _cells.GetCell(2, i).Style = StyleConstants.Success;

                            _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                            _cells.GetCell(ResultColumn03, i).Value = "ACTUALIZADO ";
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(2, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:ModifyWorkRequestMnttoList()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:UpdateReferenceCodesList()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnCleanSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
            _cells.ClearTableRange(TableName02);
            _cells.ClearTableRange(TableName03);
            //_cells.ClearTableRange(TableName04);
            _cells.ClearTableRange(TableNameM01);
            _cells.ClearTableRange(TableNameM02);
            _cells.ClearTableRange(TableNameM03);
            //_cells.ClearTableRange(TableNameM04);
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
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

                var searchCriteriaList =
                    WorkRequestActions.SearchFieldCriteriaType.GetSearchFieldCriteriaTypes()
                        .Select(g => g.Value)
                        .ToList();
                var workGroupList = GroupConstants.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = WrStatusList.GetStatusNames();
                statusList.Add(WrStatusList.Uncompleted);

                var dateCriteriaList =
                    WorkRequestActions.SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").AddComment("--ÁREA/SUPERINTENDENCIA--\n" +
                                                "INST: IMIS, MINA\n" +
                                                "MDC: FFCC, PBV, PTAS\n" +
                                                "MNTTO: MINA\n" +
                                                "SOP: ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A3").Comment.Shape.TextFrame.AutoSize = true;
                _cells.GetCell("A3").Value = WorkRequestActions.SearchFieldCriteriaType.WorkGroup.Value;
                _cells.SetValidationList(_cells.GetCell("A3"), searchCriteriaList, ValidationSheetName, 1, false);
                _cells.SetValidationList(_cells.GetCell("B3"), workGroupList, ValidationSheetName, 2, false);
                _cells.GetCell("A4").Value = WorkRequestActions.SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A4"), ValidationSheetName, 1, false);
                _cells.GetCell("A5").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B5"), statusList, ValidationSheetName, 3, false);
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = WorkRequestActions.SearchDateCriteriaType.Raised.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 4);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) +
                                             string.Format("{0:00}", DateTime.Now.Month) +
                                             string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);


                _cells.GetRange(2, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(1, TitleRow01, ResultColumn01 - 1, TitleRow01).Style = StyleConstants.TitleRequired;
                for (var i = 2; i < ResultColumn01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1)
                        .AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat =
                    NumberFormatConstants.Text;
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

                var classificationList =
                    _eFunctions.GetItemCodes("RQCL").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(08, TitleRow01 + 1), classificationList, ValidationSheetName, 5,
                    false);

                var reqTypeItemCodeList = WoTypeMtType.GetWoTypeList();
                var requestTypeList = Utils.GetCodeList(reqTypeItemCodeList);
                _cells.SetValidationList(_cells.GetCell(09, TitleRow01 + 1), requestTypeList, ValidationSheetName, 6,
                    false);

                var usTypeCodeList =
                    _eFunctions.GetItemCodes("WS").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(10, TitleRow01 + 1), usTypeCodeList, ValidationSheetName, 7,
                    false);

                var priorityList = Utils.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                _cells.SetValidationList(_cells.GetCell(11, TitleRow01 + 1), priorityList, ValidationSheetName, 8, false);

                var regionList =
                    _eFunctions.GetItemCodes("REGN").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(12, TitleRow01 + 1), regionList, ValidationSheetName, 9, false);

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
                _cells.SetValidationList(_cells.GetCell(14, TitleRow01 + 1), requestSourceList, ValidationSheetName, 10,
                    false);

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
                _cells.GetRange(28, TitleRow01, ResultColumn01 - 1, TitleRow01).Style = StyleConstants.TitleOptional;
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

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat =
                    NumberFormatConstants.Text;
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

                _cells.GetCell("A3").AddComment("--ÁREA/SUPERINTENDENCIA--\n" +
                                                "INST: IMIS, MINA\n" +
                                                "MDC: FFCC, PBV, PTAS\n" +
                                                "MNTTO: MINA\n" +
                                                "SOP: ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A3").Comment.Shape.TextFrame.AutoSize = true;
                _cells.GetCell("A3").Value = WorkRequestActions.SearchFieldCriteriaType.WorkGroup.Value;
                _cells.SetValidationList(_cells.GetCell("A3"), ValidationSheetName, 1, false);
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 2, false);
                _cells.GetCell("A4").Value = WorkRequestActions.SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A4"), ValidationSheetName, 1);
                _cells.GetCell("A5").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B5"), ValidationSheetName, 3, false);
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = WorkRequestActions.SearchDateCriteriaType.Raised.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), ValidationSheetName, 4);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) +
                                             string.Format("{0:00}", DateTime.Now.Month) +
                                             string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, TitleRow03, ResultColumn03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;
                for (var i = 4; i < ResultColumn03; i++)
                {
                    _cells.GetCell(i, TitleRow03 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow03 - 1)
                        .AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
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

                _cells.GetRange(1, TitleRow03 + 1, ResultColumn03, TitleRow03 + 1).NumberFormat =
                    NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumn03, TitleRow03 + 1), TableName03);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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

                var searchCriteriaList =
                    WorkRequestActions.SearchFieldCriteriaType.GetSearchFieldCriteriaTypes()
                        .Select(g => g.Value)
                        .ToList();
                var workGroupList = GroupConstants.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = WrStatusList.GetStatusNames();
                statusList.Add(WrStatusList.Uncompleted);

                var dateCriteriaList =
                    WorkRequestActions.SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").AddComment("--ÁREA/SUPERINTENDENCIA--\n" +
                                                "INST: IMIS, MINA\n" +
                                                "MDC: FFCC, PBV, PTAS\n" +
                                                "MNTTO: MINA\n" +
                                                "SOP: ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A3").Comment.Shape.TextFrame.AutoSize = true;
                _cells.GetCell("A3").Value = WorkRequestActions.SearchFieldCriteriaType.WorkGroup.Value;
                _cells.SetValidationList(_cells.GetCell("A3"), searchCriteriaList, ValidationSheetName, 1, false);
                _cells.SetValidationList(_cells.GetCell("B3"), workGroupList, ValidationSheetName, 2, false);
                _cells.GetCell("A4").Value = WorkRequestActions.SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A4"), ValidationSheetName, 1);
                _cells.GetCell("A5").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B5"), statusList, ValidationSheetName, 3, false);
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = WorkRequestActions.SearchDateCriteriaType.Raised.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 4);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) +
                                             string.Format("{0:00}", DateTime.Now.Month) +
                                             string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);


                _cells.GetRange(2, TitleRow01 - 2, ResultColumnM01 - 1, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(1, TitleRow01, ResultColumnM01 - 1, TitleRow01).Style = StyleConstants.TitleRequired;
                for (var i = 2; i < ResultColumnM01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1)
                        .AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }
                _cells.GetRange(1, TitleRow01 + 1, ResultColumnM01, TitleRow01 + 1).NumberFormat =
                    NumberFormatConstants.Text;
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

                var classificationList =
                    _eFunctions.GetItemCodes("RQCL").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(08, TitleRow01 + 1), classificationList, ValidationSheetName, 5,
                    false);

                var reqTypeItemCodeList = WoTypeMtType.GetWoTypeList();
                var requestTypeList = Utils.GetCodeList(reqTypeItemCodeList);
                _cells.SetValidationList(_cells.GetCell(09, TitleRow01 + 1), requestTypeList, ValidationSheetName, 6,
                    false);

                var usTypeCodeList =
                    _eFunctions.GetItemCodes("WS").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(10, TitleRow01 + 1), usTypeCodeList, ValidationSheetName, 7,
                    false);

                var priorityList = Utils.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                _cells.SetValidationList(_cells.GetCell(11, TitleRow01 + 1), priorityList, ValidationSheetName, 8, false);

                var regionList =
                    _eFunctions.GetItemCodes("REGN").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(12, TitleRow01 + 1), regionList, ValidationSheetName, 9, false);

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
                _cells.SetValidationList(_cells.GetCell(14, TitleRow01 + 1), requestSourceList, ValidationSheetName, 10,
                    false);

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

                _cells.GetRange(1, TitleRowM02, ResultColumnM02 - 1, TitleRowM02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRowM02).Value = "REQUEST ID";
                _cells.GetCell(2, TitleRowM02).Value = "CLOSED BY";
                _cells.GetCell(3, TitleRowM02).Value = "CLOSED DATE";
                _cells.GetCell(4, TitleRowM02).Value = "CLOSED TIME";
                _cells.GetCell(4, TitleRowM02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ResultColumnM02, TitleRowM02).Value = "RESULTADO";
                _cells.GetCell(ResultColumnM02, TitleRowM02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRowM02 + 1, ResultColumnM02, TitleRowM02 + 1).NumberFormat =
                    NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowM02, ResultColumnM02, TitleRowM02 + 1), TableNameM02);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                //////CONSTRUYO LA HOJA 3 SLA MNTTO
                //// ReSharper disable once UseIndexedProperty
                //_excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                //_excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameM03;

                //_cells.GetCell("A1").Value = "CERREJÓN";
                //_cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                //_cells.MergeCells("A1", "B2");
                //_cells.GetCell("C1").Value = "SERVICE LEVEL AGREEMENT WORK REQUEST - ELLIPSE 8";
                //_cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                //_cells.MergeCells("C1", "J2");

                //_cells.GetCell("K1").Value = "OBLIGATORIO";
                //_cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                //_cells.GetCell("K2").Value = "OPCIONAL";
                //_cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                //_cells.GetCell("K3").Value = "INFORMATIVO";
                //_cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                //_cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                //_cells.GetCell("K4").Style = StyleConstants.TitleAction;
                //_cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                //_cells.GetCell("K5").Style = StyleConstants.TitleAdditional;

                //_cells.GetCell("A3").AddComment("--ÁREA/SUPERINTENDENCIA--\n" +
                //     "INST: IMIS, MINA\n" +
                //     "MDC: FFCC, PBV, PTAS\n" +
                //     "MNTTO: MINA\n" +
                //     "SOP: ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                //_cells.GetCell("A3").Comment.Shape.TextFrame.AutoSize = true;
                //_cells.GetCell("A3").Value = WorkRequestActions.SearchFieldCriteriaType.WorkGroup.Value;
                //_cells.SetValidationList(_cells.GetCell("A3"), ValidationSheetName, 1, false);
                //_cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 2, false);
                //_cells.GetCell("A4").Value = WorkRequestActions.SearchFieldCriteriaType.EquipmentReference.Value;
                //_cells.SetValidationList(_cells.GetCell("A4"), ValidationSheetName, 1);
                //_cells.GetCell("A5").Value = "STATUS";
                //_cells.SetValidationList(_cells.GetCell("B5"), ValidationSheetName, 3, false);
                //_cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                //_cells.GetCell("C3").Value = "FECHA";
                //_cells.GetCell("D3").Value = WorkRequestActions.SearchDateCriteriaType.Raised.Value;
                //_cells.SetValidationList(_cells.GetCell("D3"), ValidationSheetName, 4);
                //_cells.GetCell("C4").Value = "DESDE";
                //_cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                //_cells.GetCell("D4").AddComment("YYYYMMDD");
                //_cells.GetCell("C5").Value = "HASTA";
                //_cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                //_cells.GetCell("D5").AddComment("YYYYMMDD");
                //_cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                //_cells.GetRange(1, TitleRow03, ResultColumnM03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;
                //for (var i = 4; i < ResultColumnM03; i++)
                //{
                //    _cells.GetCell(i, TitleRow03 - 1).Style = StyleConstants.ItalicSmall;
                //    _cells.GetCell(i, TitleRow03 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                //    _cells.GetCell(i, TitleRow03 - 1).Value = "true";
                //}
                //_cells.GetRange(1, TitleRow03, ResultColumnM03, TitleRow03).Style = StyleConstants.TitleOptional;
                //_cells.GetCell(02, TitleRow03).Style = StyleConstants.TitleRequired;
                //_cells.GetCell(03, TitleRow03).Style = StyleConstants.TitleInformation;

                //_cells.GetCell(01, TitleRow03).Value = "WORKGROUP";
                //_cells.GetCell(02, TitleRow03).Value = "REQUEST ID";
                //_cells.GetCell(03, TitleRow03).Value = "DESCRIPTION";

                //_cells.GetCell(04, TitleRow01).Value = "SL_AGREEMENT";
                //_cells.GetCell(05, TitleRow01).Value = "SLA_FAILURE_CODE";
                //_cells.GetCell(06, TitleRow01).Value = "SLA_START_DATE";
                //_cells.GetCell(07, TitleRow01).Value = "SLA_START_TIME";
                //_cells.GetCell(08, TitleRow01).Value = "SLA_DUE_DATE";
                //_cells.GetCell(09, TitleRow01).Value = "SLA_DUE_TIME";
                //_cells.GetCell(10, TitleRow01).Value = "SLA_DUE_DAYS";
                //_cells.GetCell(11, TitleRow01).Value = "SLA_WARN_DATE";
                //_cells.GetCell(12, TitleRow01).Value = "SLA_WARN_TIME";
                //_cells.GetCell(13, TitleRow01).Value = "SLA_WARN_DAYS";

                //_cells.GetCell(ResultColumnM03, TitleRow03).Value = "RESULTADO";
                //_cells.GetCell(ResultColumnM03, TitleRow03).Style = StyleConstants.TitleResult;

                //_cells.GetRange(1, TitleRow03 + 1, ResultColumnM03, TitleRow03 + 1).NumberFormat = NumberFormatConstants.Text;
                //_cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumnM03, TitleRow03 + 1), TableNameM03);
                //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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

            var searchCriteriaList = WorkRequestActions.SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = WorkRequestActions.SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A3").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            var listwr = WorkRequestActions.FetchWorkRequest(_eFunctions, searchCriteriaKey1, searchCriteriaValue1,
                searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
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
                    Debugger.LogError("RibbonEllipse.cs:ReviewWorkRequestList()", ex.Message);
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
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            var searchCriteriaList = WorkRequestActions.SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = WorkRequestActions.SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A3").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            var listwr = WorkRequestActions.FetchWorkRequest(_eFunctions, searchCriteriaKey1, searchCriteriaValue1,
                searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
            var i = TitleRowM01 + 1;
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

                    var wrRefCodes = WorkRequestReferenceCodesActions.GetWorkRequestReferenceCodes(_eFunctions,
                        urlService, opContext, wr.requestId);
                    //REFERENCE CODES    
                    _cells.GetCell(28, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Header;
                    _cells.GetCell(29, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Body;
                    _cells.GetCell(29, 1).WrapText = false;
                    _cells.GetCell(30, i).Value = "'" + wrRefCodes.WorkOrderOrigen;
                    _cells.GetCell(31, i).Value = "'" + wrRefCodes.StockCode1;
                    _cells.GetCell(32, i).Value = "'" + wrRefCodes.StockCode1Qty;
                    _cells.GetCell(33, i).Value = "'" + wrRefCodes.StockCode2;
                    _cells.GetCell(34, i).Value = "'" + wrRefCodes.StockCode2Qty;
                    _cells.GetCell(35, i).Value = "'" + wrRefCodes.StockCode3;
                    _cells.GetCell(36, i).Value = "'" + wrRefCodes.StockCode3Qty;
                    _cells.GetCell(37, i).Value = "'" + wrRefCodes.StockCode4;
                    _cells.GetCell(38, i).Value = "'" + wrRefCodes.StockCode4Qty;
                    _cells.GetCell(39, i).Value = "'" + wrRefCodes.StockCode5;
                    _cells.GetCell(40, i).Value = "'" + wrRefCodes.StockCode5Qty;
                    _cells.GetCell(41, i).Value = "'" + wrRefCodes.HorasHombre;
                    _cells.GetCell(42, i).Value = "'" + wrRefCodes.DuracionTarea;
                    _cells.GetCell(43, i).Value = "'" + wrRefCodes.EquipoDetenido;
                    _cells.GetCell(44, i).Value = "'" + wrRefCodes.RaisedReprogramada;
                    _cells.GetCell(45, i).Value = "'" + wrRefCodes.CambioHora;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWorkRequestList()", ex.Message);
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
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(ReReviewWorkRequest);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNamePfc01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(ReReviewWorkRequestPfc);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void ReReviewWorkRequestPfc()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNamePfc01, ResultColumnPfc01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var i = TitleRowPfc01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);
                    if (wr == null || wr.requestId == null)
                        throw new Exception("WORK REQUEST NO ENCONTRADO");
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.requestId;
                    _cells.GetCell(02, i).Value = "'" + wr.requestIdDescription1;
                    _cells.GetCell(03, i).Value = "'" + wr.requestIdDescription2;

                    //WORK                        
                    _cells.GetCell(04, i).Value = "'" + wr.employee;
                    _cells.GetCell(05, i).Value = "'" + wr.priorityCode;
                    _cells.GetCell(06, i).Value = "'" + wr.contactId;
                    _cells.GetCell(07, i).Value = "'" + wr.sourceReference;
                    _cells.GetCell(08, i).Value = "'" + wr.raisedDate;
                    _cells.GetCell(09, i).Value = "'" + wr.ServiceLevelAgreement.ServiceLevel;

                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnPfc01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWorkRequestList()", ex.Message);
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

        private void ReReviewWorkRequest()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var i = TitleRow01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);
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
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWorkRequestList()", ex.Message);
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
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNameM01, ResultColumnM01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            var i = TitleRow01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);

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
                    var wrRefCodes = WorkRequestReferenceCodesActions.GetWorkRequestReferenceCodes(_eFunctions,
                        urlService, opContext, wr.requestId);
                    //REFERENCE CODES    
                    _cells.GetCell(28, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Header;
                    _cells.GetCell(29, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Body;
                    _cells.GetCell(29, i).WrapText = false;
                    _cells.GetCell(30, i).Value = "'" + wrRefCodes.WorkOrderOrigen;
                    _cells.GetCell(31, i).Value = "'" + wrRefCodes.StockCode1;
                    _cells.GetCell(32, i).Value = "'" + wrRefCodes.StockCode1Qty;
                    _cells.GetCell(33, i).Value = "'" + wrRefCodes.StockCode2;
                    _cells.GetCell(34, i).Value = "'" + wrRefCodes.StockCode2Qty;
                    _cells.GetCell(35, i).Value = "'" + wrRefCodes.StockCode3;
                    _cells.GetCell(36, i).Value = "'" + wrRefCodes.StockCode3Qty;
                    _cells.GetCell(37, i).Value = "'" + wrRefCodes.StockCode4;
                    _cells.GetCell(38, i).Value = "'" + wrRefCodes.StockCode4Qty;
                    _cells.GetCell(39, i).Value = "'" + wrRefCodes.StockCode5;
                    _cells.GetCell(40, i).Value = "'" + wrRefCodes.StockCode5Qty;
                    _cells.GetCell(41, i).Value = "'" + wrRefCodes.HorasHombre;
                    _cells.GetCell(42, i).Value = "'" + wrRefCodes.DuracionTarea;
                    _cells.GetCell(43, i).Value = "'" + wrRefCodes.EquipoDetenido;
                    _cells.GetCell(44, i).Value = "'" + wrRefCodes.RaisedReprogramada;
                    _cells.GetCell(45, i).Value = "'" + wrRefCodes.CambioHora;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWorkRequestList()", ex.Message);
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

        private void ReviewReferenceCodesList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();
            _cells.ClearTableRange(TableName03);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            var searchCriteriaList = WorkRequestActions.SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = WorkRequestActions.SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A3").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D2").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            var listwr = WorkRequestActions.FetchWorkRequest(_eFunctions, searchCriteriaKey1, searchCriteriaValue1,
                searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
            var i = TitleRow03 + 1;

            foreach (var wr in listwr)
            {
                try
                {
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + wr.requestIdDescription1 + " " + wr.requestIdDescription2;

                    var wrRefCodes = WorkRequestReferenceCodesActions.GetWorkRequestReferenceCodes(_eFunctions,
                        urlService, opContext, wr.requestId);
                    _cells.GetCell(04, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Header;
                    _cells.GetCell(05, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Body;
                    _cells.GetCell(05, i).WrapText = false;
                    _cells.GetCell(06, i).Value = "'" + wrRefCodes.WorkOrderOrigen;
                    _cells.GetCell(07, i).Value = "'" + wrRefCodes.StockCode1;
                    _cells.GetCell(08, i).Value = "'" + wrRefCodes.StockCode1Qty;
                    _cells.GetCell(09, i).Value = "'" + wrRefCodes.StockCode2;
                    _cells.GetCell(10, i).Value = "'" + wrRefCodes.StockCode2Qty;
                    _cells.GetCell(11, i).Value = "'" + wrRefCodes.StockCode3;
                    _cells.GetCell(12, i).Value = "'" + wrRefCodes.StockCode3Qty;
                    _cells.GetCell(13, i).Value = "'" + wrRefCodes.StockCode4;
                    _cells.GetCell(14, i).Value = "'" + wrRefCodes.StockCode4Qty;
                    _cells.GetCell(15, i).Value = "'" + wrRefCodes.StockCode5;
                    _cells.GetCell(16, i).Value = "'" + wrRefCodes.StockCode5Qty;
                    _cells.GetCell(17, i).Value = "'" + wrRefCodes.HorasHombre;
                    _cells.GetCell(18, i).Value = "'" + wrRefCodes.HorasQty;
                    _cells.GetCell(19, i).Value = "'" + wrRefCodes.DuracionTarea;
                    _cells.GetCell(20, i).Value = "'" + wrRefCodes.EquipoDetenido;
                    _cells.GetCell(21, i).Value = "'" + wrRefCodes.RaisedReprogramada;
                    _cells.GetCell(22, i).Value = "'" + wrRefCodes.CambioHora;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewReferenceCodesList()", ex.Message);
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
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ReReviewReferenceCodesList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName03, ResultColumn03);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            var i = TitleRow03 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wr = WorkRequestActions.FetchWorkRequest(_eFunctions, requestId);

                    if (wr == null || wr.requestId == null)
                        throw new Exception("WORK REQUEST NO ENCONTRADO");
                    //GENERAL
                    _cells.GetCell(01, i).Value = "'" + wr.workGroup;
                    _cells.GetCell(02, i).Value = "'" + wr.requestId;
                    _cells.GetCell(03, i).Value = "'" + wr.requestIdDescription1 + " " + wr.requestIdDescription2;
                    var wrRefCodes = WorkRequestReferenceCodesActions.GetWorkRequestReferenceCodes(_eFunctions,
                        urlService, opContext, wr.requestId);
                    _cells.GetCell(04, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Header;
                    _cells.GetCell(05, i).Value = "'" + wr.GetExtendedDescription(urlService, opContext).Body;
                    _cells.GetCell(05, i).WrapText = false;
                    _cells.GetCell(06, i).Value = "'" + wrRefCodes.WorkOrderOrigen;
                    _cells.GetCell(07, i).Value = "'" + wrRefCodes.StockCode1;
                    _cells.GetCell(08, i).Value = "'" + wrRefCodes.StockCode1Qty;
                    _cells.GetCell(09, i).Value = "'" + wrRefCodes.StockCode2;
                    _cells.GetCell(10, i).Value = "'" + wrRefCodes.StockCode2Qty;
                    _cells.GetCell(11, i).Value = "'" + wrRefCodes.StockCode3;
                    _cells.GetCell(12, i).Value = "'" + wrRefCodes.StockCode3Qty;
                    _cells.GetCell(13, i).Value = "'" + wrRefCodes.StockCode4;
                    _cells.GetCell(14, i).Value = "'" + wrRefCodes.StockCode4Qty;
                    _cells.GetCell(15, i).Value = "'" + wrRefCodes.StockCode5;
                    _cells.GetCell(16, i).Value = "'" + wrRefCodes.StockCode5Qty;
                    _cells.GetCell(17, i).Value = "'" + wrRefCodes.HorasHombre;
                    _cells.GetCell(18, i).Value = "'" + wrRefCodes.HorasQty;
                    _cells.GetCell(19, i).Value = "'" + wrRefCodes.DuracionTarea;
                    _cells.GetCell(20, i).Value = "'" + wrRefCodes.EquipoDetenido;
                    _cells.GetCell(21, i).Value = "'" + wrRefCodes.RaisedReprogramada;
                    _cells.GetCell(22, i).Value = "'" + wrRefCodes.CambioHora;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewReferenceCodesList()", ex.Message);
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

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow01 + 1;

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
                        classification = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value)),
                        requestType = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)),
                        userStatus = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value)),
                        priorityCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value)),
                        region = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value)),
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
                            FailureCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(29, i).Value)),
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
                    Debugger.LogError("RibbonEllipse.cs:CreateWorkRequestList()", ex.Message);
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

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow01 + 1;

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
                        classification = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value)),
                        requestType = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)),
                        userStatus = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value)),
                        priorityCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value)),
                        region = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value)),
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
                        standardJobDistrict = _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value)
                    };
                    var header = _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value);
                    var body = _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value);
                    wr.SetExtendedDescription(header, body);
                    var wrRefCodes = new WorkRequestReferenceCodes
                    {
                        WorkOrderOrigen = _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value),
                        StockCode1 = _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value),
                        StockCode1Qty = _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value),
                        StockCode2 = _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value),
                        StockCode2Qty = _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value),
                        StockCode3 = _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value),
                        StockCode3Qty = _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value),
                        StockCode4 = _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value),
                        StockCode4Qty = _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value),
                        StockCode5 = _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value),
                        StockCode5Qty = _cells.GetEmptyIfNull(_cells.GetCell(40, i).Value),
                        HorasHombre = _cells.GetEmptyIfNull(_cells.GetCell(41, i).Value),
                        DuracionTarea = _cells.GetEmptyIfNull(_cells.GetCell(42, i).Value),
                        EquipoDetenido = _cells.GetEmptyIfNull(_cells.GetCell(43, i).Value),
                        RaisedReprogramada = _cells.GetEmptyIfNull(_cells.GetCell(44, i).Value),
                        CambioHora = _cells.GetEmptyIfNull(_cells.GetCell(45, i).Value)
                    };

                    var replySheet = WorkRequestActions.CreateWorkRequest(urlService, opContext, wr);
                    var requestId = replySheet.requestId;
                    if (string.IsNullOrWhiteSpace(replySheet.requestId))
                        throw new Exception("No se ha podido crear el WorkRequest");
                    var errorList = "";
                    var replyExtended = WorkRequestActions.UpdateWorkRequestExtendedDescription(urlService, opContext,
                        requestId, wr.GetExtendedDescription(urlService, opContext));
                    if (replyExtended != null && replyExtended.Errors != null && replyExtended.Errors.Length > 0)
                        errorList = replyExtended.Errors.Aggregate(errorList, (current, error) => current + ("\nError: " + error));

                    var replyRefCode = WorkRequestReferenceCodesActions.ModifyReferenceCodes(_eFunctions, urlService,
                        opContext, requestId, wrRefCodes);
                    if (replyRefCode != null && replyRefCode.Errors != null && replyRefCode.Errors.Length > 0)
                        errorList = replyExtended.Errors.Aggregate(errorList, (current, error) => current + ("\nError: " + error));

                    if (!string.IsNullOrWhiteSpace(errorList))
                    {
                        _cells.GetCell(2, i).Value = "'" + requestId;
                        _cells.GetCell(2, i).Style = StyleConstants.Success;

                        _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Warning;
                        _cells.GetCell(ResultColumnM01, i).Value = "CREADO " + requestId + errorList;
                    }
                    else
                    {
                        _cells.GetCell(2, i).Value = "'" + requestId;
                        _cells.GetCell(2, i).Style = StyleConstants.Success;

                        _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumnM01, i).Value = "CREADO " + requestId;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWorkRequestMnttoList()", ex.Message);
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
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(ModifyWorkRequest);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNamePfc01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(ModifyWorkRequestPfc);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel no tiene el formato requerido");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateWorkRequest()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void ModifyWorkRequestPfc()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNamePfc01, ResultColumnPfc01);

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowPfc01 + 1;

            var employee = _frmAuth.EllipseUser;
            var todayDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup = "PLANFC",
                        requestIdDescription1 = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        requestIdDescription2 = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        equipmentNo = "FERROCARRIL",
                        employee = string.IsNullOrEmpty(_cells.GetEmptyIfNull(_cells.GetCell(4, i).Value))
                                ? employee
                                : _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value),
                        classification = "SS",
                        requestType = "ES",
                        priorityCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(5, i).Value)),
                        contactId = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),
                        sourceReference = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        raisedDate = string.IsNullOrWhiteSpace(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value))
                                    ? todayDate
                                    : _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        ServiceLevelAgreement =
                        {
                            ServiceLevel = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)),
                            StartDate = todayDate
                        }
                    };

                    WorkRequestActions.ModifyWorkRequest(urlService, opSheet, wr);
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnPfc01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnPfc01, i).Value = "ACTUALIZADO";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnPfc01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnPfc01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ModifyWorkRequestList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ModifyWorkRequest()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup =
                            Utils.IsTrue(_cells.GetCell(1, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value)
                                : null,
                        requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        requestIdDescription1 =
                            Utils.IsTrue(_cells.GetCell(4, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value)
                                : null,
                        requestIdDescription2 =
                            Utils.IsTrue(_cells.GetCell(5, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value)
                                : null,
                        equipmentNo =
                            Utils.IsTrue(_cells.GetCell(6, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value)
                                : null,
                        employee =
                            Utils.IsTrue(_cells.GetCell(7, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value)
                                : null,
                        classification =
                            Utils.IsTrue(_cells.GetCell(8, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value))
                                : null,
                        requestType =
                            Utils.IsTrue(_cells.GetCell(9, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value))
                                : null,
                        userStatus =
                            Utils.IsTrue(_cells.GetCell(10, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value))
                                : null,
                        priorityCode =
                            Utils.IsTrue(_cells.GetCell(11, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value))
                                : null,
                        region =
                            Utils.IsTrue(_cells.GetCell(12, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value))
                                : null,
                        contactId =
                            Utils.IsTrue(_cells.GetCell(13, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value)
                                : null,
                        source =
                            Utils.IsTrue(_cells.GetCell(14, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)
                                : null,
                        sourceReference =
                            Utils.IsTrue(_cells.GetCell(15, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value)
                                : null,
                        requiredByDate =
                            Utils.IsTrue(_cells.GetCell(16, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value)
                                : null,
                        requiredByTime =
                            Utils.IsTrue(_cells.GetCell(17, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value)
                                : null,
                        raisedUser =
                            Utils.IsTrue(_cells.GetCell(18, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value)
                                : null,
                        raisedDate =
                            Utils.IsTrue(_cells.GetCell(19, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value)
                                : null,
                        raisedTime =
                            Utils.IsTrue(_cells.GetCell(20, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value)
                                : null,
                        closedBy =
                            Utils.IsTrue(_cells.GetCell(21, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value)
                                : null,
                        closedDate =
                            Utils.IsTrue(_cells.GetCell(22, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value)
                                : null,
                        assignPerson =
                            Utils.IsTrue(_cells.GetCell(23, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value)
                                : null,
                        ownerId =
                            Utils.IsTrue(_cells.GetCell(24, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value)
                                : null,
                        estimateNo =
                            Utils.IsTrue(_cells.GetCell(25, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value)
                                : null,
                        standardJob =
                            Utils.IsTrue(_cells.GetCell(26, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value)
                                : null,
                        standardJobDistrict =
                            Utils.IsTrue(_cells.GetCell(27, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value)
                                : null,
                        ServiceLevelAgreement =
                        {
                            ServiceLevel =
                                Utils.IsTrue(_cells.GetCell(28, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value)
                                    : null,
                            FailureCode =
                                Utils.IsTrue(_cells.GetCell(29, validationRow).Value)
                                    ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(29, i).Value))
                                    : null,
                            StartDate =
                                Utils.IsTrue(_cells.GetCell(30, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value)
                                    : null,
                            StartTime =
                                Utils.IsTrue(_cells.GetCell(31, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value)
                                    : null,
                            DueDate =
                                Utils.IsTrue(_cells.GetCell(32, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value)
                                    : null,
                            DueTime =
                                Utils.IsTrue(_cells.GetCell(33, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value)
                                    : null,
                            DueDays =
                                Utils.IsTrue(_cells.GetCell(34, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value)
                                    : null,
                            WarnDate =
                                Utils.IsTrue(_cells.GetCell(35, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value)
                                    : null,
                            WarnTime =
                                Utils.IsTrue(_cells.GetCell(36, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value)
                                    : null,
                            WarnDays =
                                Utils.IsTrue(_cells.GetCell(37, validationRow).Value)
                                    ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value)
                                    : null
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
                    Debugger.LogError("RibbonEllipse.cs:ModifyWorkRequestList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ModifyWorkRequestMnttoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNameM01, ResultColumnM01);

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup =
                            Utils.IsTrue(_cells.GetCell(1, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value)
                                : null,
                        requestId =
                            Utils.IsTrue(_cells.GetCell(2, validationRow).Value)
                                ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value)
                                : null,
                        requestIdDescription1 =
                            Utils.IsTrue(_cells.GetCell(4, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value)
                                : null,
                        requestIdDescription2 =
                            Utils.IsTrue(_cells.GetCell(5, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value)
                                : null,
                        equipmentNo =
                            Utils.IsTrue(_cells.GetCell(6, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value)
                                : null,
                        employee =
                            Utils.IsTrue(_cells.GetCell(7, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value)
                                : null,
                        classification =
                            Utils.IsTrue(_cells.GetCell(8, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value))
                                : null,
                        requestType =
                            Utils.IsTrue(_cells.GetCell(9, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value))
                                : null,
                        userStatus =
                            Utils.IsTrue(_cells.GetCell(10, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value))
                                : null,
                        priorityCode =
                            Utils.IsTrue(_cells.GetCell(11, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value))
                                : null,
                        region =
                            Utils.IsTrue(_cells.GetCell(12, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value))
                                : null,
                        contactId =
                            Utils.IsTrue(_cells.GetCell(13, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value)
                                : null,
                        source =
                            Utils.IsTrue(_cells.GetCell(14, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)
                                : null,
                        sourceReference =
                            Utils.IsTrue(_cells.GetCell(15, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value)
                                : null,
                        requiredByDate =
                            Utils.IsTrue(_cells.GetCell(16, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value)
                                : null,
                        requiredByTime =
                            Utils.IsTrue(_cells.GetCell(17, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value)
                                : null,
                        raisedUser =
                            Utils.IsTrue(_cells.GetCell(18, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value)
                                : null,
                        raisedDate =
                            Utils.IsTrue(_cells.GetCell(19, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value)
                                : null,
                        raisedTime =
                            Utils.IsTrue(_cells.GetCell(20, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value)
                                : null,
                        closedBy =
                            Utils.IsTrue(_cells.GetCell(21, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value)
                                : null,
                        closedDate =
                            Utils.IsTrue(_cells.GetCell(22, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value)
                                : null,
                        assignPerson =
                            Utils.IsTrue(_cells.GetCell(23, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value)
                                : null,
                        ownerId =
                            Utils.IsTrue(_cells.GetCell(24, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value)
                                : null,
                        estimateNo =
                            Utils.IsTrue(_cells.GetCell(25, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value)
                                : null,
                        standardJob =
                            Utils.IsTrue(_cells.GetCell(26, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value)
                                : null,
                        standardJobDistrict =
                            Utils.IsTrue(_cells.GetCell(27, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value)
                                : null
                    };

                    var header = Utils.IsTrue(_cells.GetCell(28, validationRow).Value)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value)
                        : null;
                    var body = Utils.IsTrue(_cells.GetCell(29, validationRow).Value)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value)
                        : null;
                    wr.SetExtendedDescription(header, body);
                    var wrRefCodes = new WorkRequestReferenceCodes
                    {
                        WorkOrderOrigen =
                            Utils.IsTrue(_cells.GetCell(30, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value)
                                : null,
                        StockCode1 =
                            Utils.IsTrue(_cells.GetCell(31, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value)
                                : null,
                        StockCode1Qty =
                            Utils.IsTrue(_cells.GetCell(32, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value)
                                : null,
                        StockCode2 =
                            Utils.IsTrue(_cells.GetCell(33, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value)
                                : null,
                        StockCode2Qty =
                            Utils.IsTrue(_cells.GetCell(34, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value)
                                : null,
                        StockCode3 =
                            Utils.IsTrue(_cells.GetCell(35, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value)
                                : null,
                        StockCode3Qty =
                            Utils.IsTrue(_cells.GetCell(36, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value)
                                : null,
                        StockCode4 =
                            Utils.IsTrue(_cells.GetCell(37, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value)
                                : null,
                        StockCode4Qty =
                            Utils.IsTrue(_cells.GetCell(38, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value)
                                : null,
                        StockCode5 =
                            Utils.IsTrue(_cells.GetCell(39, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value)
                                : null,
                        StockCode5Qty =
                            Utils.IsTrue(_cells.GetCell(40, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(40, i).Value)
                                : null,
                        HorasHombre =
                            Utils.IsTrue(_cells.GetCell(41, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(41, i).Value)
                                : null,
                        DuracionTarea =
                            Utils.IsTrue(_cells.GetCell(42, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(42, i).Value)
                                : null,
                        EquipoDetenido =
                            Utils.IsTrue(_cells.GetCell(43, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(43, i).Value)
                                : null,
                        RaisedReprogramada =
                            Utils.IsTrue(_cells.GetCell(44, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(44, i).Value)
                                : null,
                        CambioHora =
                            Utils.IsTrue(_cells.GetCell(45, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(45, i).Value)
                                : null
                    };

                    var replySheet = WorkRequestActions.ModifyWorkRequest(urlService, opContext, wr);
                    var requestId = replySheet.requestId;
                    if (string.IsNullOrWhiteSpace(replySheet.requestId))
                        throw new Exception("No se ha podido modificar el WorkRequest");
                    var errorList = "";
                    var replyExtended = WorkRequestActions.UpdateWorkRequestExtendedDescription(urlService, opContext,
                        requestId, wr.GetExtendedDescription(urlService, opContext));
                    if (replyExtended != null && replyExtended.Errors != null && replyExtended.Errors.Length > 0)
                        foreach (var error in replyExtended.Errors)
                            errorList += "\nError: " + error;

                    var replyRefCode = WorkRequestReferenceCodesActions.ModifyReferenceCodes(_eFunctions, urlService,
                        opContext, requestId, wrRefCodes);
                    if (replyRefCode != null && replyRefCode.Errors != null && replyRefCode.Errors.Length > 0)
                        foreach (var error in replyExtended.Errors)
                            errorList += "\nError: " + error;

                    if (!string.IsNullOrWhiteSpace(errorList))
                    {
                        _cells.GetCell(2, i).Value = "'" + requestId;
                        _cells.GetCell(2, i).Style = StyleConstants.Success;

                        _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Warning;
                        _cells.GetCell(ResultColumnM01, i).Value = "ACTUALIZADO " + errorList;
                    }
                    else
                    {
                        _cells.GetCell(2, i).Value = "'" + requestId;
                        _cells.GetCell(2, i).Style = StyleConstants.Success;

                        _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumnM01, i).Value = "ACTUALIZADO ";
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ModifyWorkRequestMnttoList()", ex.Message);
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
                returnWarnings = Debugger.DebugWarnings,
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
                    Debugger.LogError("RibbonEllipse.cs:CloseWorkRequestList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void CloseWorkRequestMnttoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameM02, ResultColumnM02);

            var i = TitleRowM02 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
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
                    _cells.GetCell(ResultColumnM02, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnM02, i).Value = "CERRADO";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CloseWorkRequestMnttoList()", ex.Message);
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
                returnWarnings = Debugger.DebugWarnings,
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
                    Debugger.LogError("RibbonEllipse.cs:ReOpenWorkRequestList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ReOpenWorkRequestMnttoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameM02, ResultColumnM02);

            var i = TitleRowM02 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);

                    WorkRequestActions.ReOpenWorkRequest(urlService, opSheet, requestId);
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnM02, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnM02, i).Value = "REABIERTA";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReOpenWorkRequestMnttoList()", ex.Message);
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
                returnWarnings = Debugger.DebugWarnings,
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
                    Debugger.LogError("RibbonEllipse.cs:DeleteWorkRequestList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void DeleteWorkRequestMnttoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameM01, ResultColumnM01);

            var i = TitleRowM01 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);

                    WorkRequestActions.DeleteWorkRequest(urlService, opSheet, requestId);
                    _cells.GetCell(2, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnM01, i).Value = "ELIMINADO";
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnM01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:DeleteWorkRequestList()", ex.Message);
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
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var serviceLevelAgreement = new ServiceLevelAgreement
                    {
                        ServiceLevel =
                            Utils.IsTrue(_cells.GetCell(28, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value)
                                : null,
                        FailureCode =
                            Utils.IsTrue(_cells.GetCell(29, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(29, i).Value))
                                : null,
                        StartDate =
                            Utils.IsTrue(_cells.GetCell(30, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value)
                                : null,
                        StartTime =
                            Utils.IsTrue(_cells.GetCell(31, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value)
                                : null,
                        DueDate =
                            Utils.IsTrue(_cells.GetCell(32, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value)
                                : null,
                        DueTime =
                            Utils.IsTrue(_cells.GetCell(33, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value)
                                : null,
                        DueDays =
                            Utils.IsTrue(_cells.GetCell(34, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value)
                                : null,
                        WarnDate =
                            Utils.IsTrue(_cells.GetCell(35, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value)
                                : null,
                        WarnTime =
                            Utils.IsTrue(_cells.GetCell(36, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value)
                                : null,
                        WarnDays =
                            Utils.IsTrue(_cells.GetCell(37, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value)
                                : null
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
                    Debugger.LogError("RibbonEllipse.cs:SetSlaList()", ex.Message);
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
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var requestId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var serviceLevelAgreement = new ServiceLevelAgreement
                    {
                        ServiceLevel =
                            Utils.IsTrue(_cells.GetCell(28, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value)
                                : null,
                        FailureCode =
                            Utils.IsTrue(_cells.GetCell(29, validationRow).Value)
                                ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(29, i).Value))
                                : null,
                        StartDate =
                            Utils.IsTrue(_cells.GetCell(30, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value)
                                : null,
                        StartTime =
                            Utils.IsTrue(_cells.GetCell(31, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value)
                                : null,
                        DueDate =
                            Utils.IsTrue(_cells.GetCell(32, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value)
                                : null,
                        DueTime =
                            Utils.IsTrue(_cells.GetCell(33, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value)
                                : null,
                        DueDays =
                            Utils.IsTrue(_cells.GetCell(34, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value)
                                : null,
                        WarnDate =
                            Utils.IsTrue(_cells.GetCell(35, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value)
                                : null,
                        WarnTime =
                            Utils.IsTrue(_cells.GetCell(36, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value)
                                : null,
                        WarnDays =
                            Utils.IsTrue(_cells.GetCell(37, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value)
                                : null
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
                    Debugger.LogError("RibbonEllipse.cs:ResetSlaList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }

            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnFormatFcVagones_Click(object sender, RibbonControlEventArgs e)
        {
            FormatFcVagones();
        }

        private void FormatFcVagones()
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
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameV01;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK REQUEST REGISTRO DE FALLAS VAGONES - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                //_cells.GetCell("A3").Value = "DESDE";
                //_cells.GetCell("B3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                //_cells.GetCell("B3").AddComment("YYYYMMDD");
                //_cells.GetCell("A4").Value = "HASTA";
                //_cells.GetCell("B4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                //_cells.GetCell("B4").AddComment("YYYYMMDD");
                //_cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                //_cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(01, TitleRowV01, ResultColumnV01 - 1, TitleRowV01).Style = StyleConstants.TitleRequired;
                //GENERAL
                _cells.GetCell(01, TitleRowV01).Value = "DESCRIPCIÓN";
                _cells.GetCell(02, TitleRowV01).Value = "ACCIÓN A REALIZAR";
                _cells.GetCell(03, TitleRowV01).Value = "EQUIPO";
                _cells.GetCell(03, TitleRowV01).AddComment("110XXXX (Ej. Vagón 300 es 1100300. Vagón 1040 es 1101040)");
                _cells.GetCell(04, TitleRowV01).Value = "EMPLEADO";
                _cells.GetCell(04, TitleRowV01)
                    .AddComment("Si no se digita usará el usuario de autenticación de Ellipse");
                _cells.GetCell(04, TitleRowV01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(05, TitleRowV01).Value = "CLASIFICACIÓN";
                _cells.GetCell(06, TitleRowV01).Value = "PRIORIDAD";
                _cells.GetCell(07, TitleRowV01).Value = "SL AGREEMENT";
                _cells.GetCell(08, TitleRowV01).Value = "SLA COD FALLA";
                _cells.GetCell(09, TitleRowV01).Value = "SLA FECHA INICIO";
                _cells.GetCell(09, TitleRowV01).AddComment("YYYYMMDD. Si no se digita usará la fecha del día de hoy");
                _cells.GetCell(09, TitleRowV01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRowV01).Value = "ASIGNADO A";
                _cells.GetCell(10, TitleRowV01).AddComment("Si no se digita usará BSOLANO de forma predeterminada");
                _cells.GetCell(10, TitleRowV01).Style = StyleConstants.TitleOptional;
                //_cells.GetCell(11, TitleRowV01).Value = "FECHA CREACIÓN";
                //_cells.GetCell(11, TitleRowV01).Style = StyleConstants.TitleInformation;

                var actionList = new List<string> { "HACER SEGUIMIENTO", "SOLICITAR A OPERACIONES REPARAR" };
                _cells.SetValidationList(_cells.GetCell(02, TitleRowV01 + 1), actionList, ValidationSheetName, 1);

                var clasificationList = new List<string>
                {
                    "ME - MECANICO",
                    "ES - ESTRUCTURAL",
                    "NE - NEUMATICO",
                    "ET - ELECTRICO"
                };
                _cells.SetValidationList(_cells.GetCell(05, TitleRowV01 + 1), clasificationList, ValidationSheetName, 2,
                    false);

                var priorityList = new List<string> { "P1 - EMERGENCIA", "P2 - ALTA", "P3 - NORMAL", "P4 - BAJA" };
                _cells.SetValidationList(_cells.GetCell(06, TitleRowV01 + 1), priorityList, ValidationSheetName, 3,
                    false);

                var agreementList = new List<string> { "1D - UN DÍA", "7D - 7 DÍAS", "14 - 14 DÍAS", "1Y - 1 AÑO" };
                _cells.SetValidationList(_cells.GetCell(07, TitleRowV01 + 1), agreementList, ValidationSheetName, 4,
                    false);

                var failureList = new List<string>
                {
                    "03 - SISTEMA DE APERTURA",
                    "07 - SISTEMA ESTRUCTURAL",
                    "04 - SISTEMA ELECTRICO",
                    "13 - SISTEMA NEUMATICO"
                };
                _cells.SetValidationList(_cells.GetCell(08, TitleRowV01 + 1), failureList, ValidationSheetName, 5, false);


                //
                _cells.GetCell(ResultColumnV01, TitleRowV01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnV01, TitleRowV01).Style = StyleConstants.TitleResult;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRowV01, ResultColumnV01, TitleRowV01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatFcVagones()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void CreateWorkRequestVagonesList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNameV01, ResultColumnV01);

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowV01 + 1;
            //default values
            var todayDate = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) +
                            string.Format("{0:00}", DateTime.Now.Day);
            //To Do change for ICARROS Group Admin
            const string assignPerson = "BSOLANO";
            var employee = _frmAuth.EllipseUser;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wr = new WorkRequest
                    {
                        workGroup = "ICARROS",
                        requestId = null,
                        requestIdDescription1 = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        requestIdDescription2 = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        equipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        employee =
                            string.IsNullOrEmpty(_cells.GetEmptyIfNull(_cells.GetCell(4, i).Value))
                                ? employee
                                : _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value),
                        classification = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(5, i).Value)),
                        priorityCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(6, i).Value)),
                        assignPerson =
                            string.IsNullOrEmpty(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value))
                                ? assignPerson
                                : _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                        ServiceLevelAgreement =
                        {
                            ServiceLevel = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(7, i).Value)),
                            FailureCode = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value)),
                            StartDate =
                                string.IsNullOrWhiteSpace(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value))
                                    ? todayDate
                                    : _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)
                        }
                    };

                    if (string.IsNullOrWhiteSpace(wr.ServiceLevelAgreement.ServiceLevel) ||
                        string.IsNullOrWhiteSpace(wr.ServiceLevelAgreement.FailureCode) ||
                        string.IsNullOrWhiteSpace(wr.ServiceLevelAgreement.StartDate))
                        throw new Exception("No se puede crear Work Request. Falta la información del Service Level");
                    var replySheet = WorkRequestActions.CreateWorkRequest(urlService, opSheet, wr);
                    var requestId = replySheet.requestId;

                    WorkRequestActions.SetWorkRequestSla(urlService, opSheet, requestId, wr.ServiceLevelAgreement);
                    _cells.GetCell(ResultColumnV01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnV01, i).Value = requestId;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(2, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnV01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnV01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWorkRequestVagonesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumnV01, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnPlanFc_Click(object sender, RibbonControlEventArgs e)
        {
            FormatPfc();
        }

        private void FormatPfc()
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
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNamePfc01;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK REQUEST REGISTRO DE FALLAS VAGONES - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(01, TitleRowPfc01, ResultColumnPfc01 - 1, TitleRowPfc01).Style =
                    StyleConstants.TitleRequired;
                //GENERAL
                _cells.GetCell(01, TitleRowPfc01).Value = "REQUEST ID";
                _cells.GetCell(02, TitleRowPfc01).Value = "DESCRIPCIÓN 1";
                _cells.GetCell(03, TitleRowPfc01).Value = "DESCRIPCIÓN 2";
                _cells.GetCell(04, TitleRowPfc01).Value = "SOLICITADO POR";
                _cells.GetCell(04, TitleRowPfc01).AddComment("Si no se digita usará el usuario de autenticación de Ellipse");
                _cells.GetCell(04, TitleRowPfc01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(05, TitleRowPfc01).Value = "PRIORIDAD";
                _cells.GetCell(06, TitleRowPfc01).Value = "ID DE SEGUIMIENTO";
                _cells.GetCell(07, TitleRowPfc01).Value = "REFERENCIA";
                _cells.GetCell(08, TitleRowPfc01).Value = "FECHA";
                _cells.GetCell(09, TitleRowPfc01).Value = "NIVEL DE SERVICIO";

                var priorityList = new List<string> { "P1 - EMERGENCIA", "P2 - ALTA", "P3 - NORMAL", "P4 - BAJA" };
                _cells.SetValidationList(_cells.GetCell(05, TitleRowPfc01 + 1), priorityList, ValidationSheetName, 1,
                    false);

                var agreementList = new List<string> { "1D - UN DÍA", "7D - 7 DÍAS", "14 - 14 DÍAS", "1Y - 1 AÑO" };
                _cells.SetValidationList(_cells.GetCell(09, TitleRowPfc01 + 1), agreementList, ValidationSheetName, 2,
                    false);

                var referenceList = new List<string>
                {
                    "CONTRATACION MAYOR",
                    "CONTRATACION DELEGADA",
                    "IMIS",
                    "VPP",
                    "CAPEX",
                    "OTRO"
                };
                _cells.SetValidationList(_cells.GetCell(07, TitleRowPfc01 + 1), referenceList, ValidationSheetName, 3,
                    false);

                //
                _cells.GetCell(ResultColumnPfc01, TitleRowPfc01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnPfc01, TitleRowPfc01).Style = StyleConstants.TitleResult;


                _cells.FormatAsTable(_cells.GetRange(1, TitleRowPfc01, ResultColumnPfc01, TitleRowPfc01 + 1),
                    TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatFcVagones()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
    }
}