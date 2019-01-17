using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using EllipseWorkOrdersClassLibrary;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using FormAuthenticate = EllipseCommonsClassLibrary.FormAuthenticate;
using EllipseStdTextClassLibrary;
using WorkOrderTaskService = EllipseWorkOrdersClassLibrary.WorkOrderTaskService;
using WorkOrderService = EllipseWorkOrdersClassLibrary.WorkOrderService;
using ResourceReqmntsService = EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using MaterialReqmntsService = EllipseWorkOrdersClassLibrary.MaterialReqmntsService;
using EquipmentReqmntsService = EllipseWorkOrdersClassLibrary.EquipmentReqmntsService;
// ReSharper disable UseIndexedProperty
// ReSharper disable UseNullPropagation

namespace EllipseWorkOrderExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Application _excelApp;

        private const string SheetName01 = "WorkOrders";
        private const string SheetName02 = "Tasks";
        private const string SheetName03 = "Requirement";
        private const string SheetName04 = "CloseWorkOrders";
        private const string SheetName05 = "CloseCommentsWorkOrders";
        private const string SheetName06 = "DurationWorkOrders";
        private const string SheetName07 = "ProgressWorkOrders";

        private const string SheetNameD01 = "WorkOrdersDetailed";
        private const string SheetNameD02 = "WOTasks";
        private const string SheetNameD03 = "WORequirements";
        private const string SheetNameD04 = "WOReferenceCodes";
        private const string SheetNameQ01 = "QualityWorkOrders";
        private const string SheetNameCc01 = "Critical Controls";

        private const int TitleRow01 = 9;
        private const int TitleRow02 = 9;
        private const int TitleRow03 = 9;
        private const int TitleRow04 = 6;
        private const int TitleRow05 = 6;
        private const int TitleRow06 = 6;
        private const int TitleRow07 = 6;

        private const int TitleRowD01 = 9;
        private const int TitleRowD02 = 6;
        private const int TitleRowD03 = 6;
        private const int TitleRowD04 = 6;
        private const int TitleRowQ01 = 7;
        private const int TitleRowCc01 = 6;

        private const int ResultColumn01 = 54;
        private const int ResultColumn02 = 24;
        private const int ResultColumn03 = 16;
        private const int ResultColumn04 = 8;
        private const int ResultColumn05 = 3;
        private const int ResultColumn06 = 8;
        private const int ResultColumn07 = 6;

        private const int ResultColumnD01 = 56;
        private const int ResultColumnD02 = 8;
        private const int ResultColumnD03 = 3;
        private const int ResultColumnD04 = 35;
        private const int ResultColumnQ01 = 36;
        private const int ResultColumnCc01 = 20;

        private const string TableName01 = "WorkOrderTable";
        private const string TableName02 = "TaskTable";
        private const string TableName03 = "RequirementTable";
        private const string TableName04 = "WorkOrderCloseTable";
        private const string TableName05 = "WorkOrderCompleteTextTable";
        private const string TableName06 = "WorkOrderDurationTable";
        private const string TableName07 = "WorkOrderProgressTable";

        private const string TableNameD01 = "WorkOrderTable";
        private const string TableNameD02 = "WorkOrderTasksTable";
        private const string TableNameD03 = "WorkOrderRequirmentsTable";
        private const string TableNameD04 = "WorkOrderReferenceCodesTable";
        private const string TableNameQ01 = "WorkOrderQualityTable";
        private const string TableNameCc01 = "CriticalControlsTable";
        private const string ValidationSheetName = "ValidationSheetWorkOrder";
        private Thread _thread;
        private bool _progressUpdate = true;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviroments = Environments.GetEnviromentList();
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
        private void btnFormatDetail_Click(object sender, RibbonControlEventArgs e)
        {
            FormatDetailed();
        }
        private void btnFormatQuality_Click(object sender, RibbonControlEventArgs e)
        {
            FormatQuality();
        }

        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReviewWoDetailedList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReReview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    _thread = new Thread(ReReviewWoDetailedList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReReviewWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnCreate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(CreateWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(CreateWoDetailedList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateWoDetailedList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnClose_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(CompleteWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CloseWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReOpen_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName04)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReOpenWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReOpenWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReviewCloseText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReviewCloseText);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewCloseText()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateCloseText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName05)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateCloseText);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateCloseText()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnDurationsReview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName06)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(GetDurationWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:GetDurationWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnDurationsAction_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName06)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ExecuteDurationWoActions);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ExecuteDurationWoActions()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReviewWorkProgress_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName07)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewWorkProgress);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWorkProgress()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdatePercentProgress_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName07)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _progressUpdate = true;
                    _thread = new Thread(UpdateWorkProgress);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateWorkProgress()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }


        private void btnUpdateUnitsProgress_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName07)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _progressUpdate = false;
                    _thread = new Thread(UpdateWorkProgress);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateWorkProgress()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateUnitsRequired_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName07)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _progressUpdate = false;
                    _thread = new Thread(UpdateRequiredProgress);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateWorkProgress()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void ReviewWorkProgress()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName07, ResultColumn07);

            var i = TitleRow07 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    //GENERAL
                    var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    string districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
                    var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, districtCode, workOrder);

                    _cells.GetCell(2, i).Value = wo.unitOfWork;
                    _cells.GetCell(3, i).Value = wo.unitsRequired;
                    _cells.GetCell(4, i).Value = wo.pcComplete;
                    _cells.GetCell(5, i).Value = wo.unitsComplete;
                    _cells.GetCell(ResultColumn07, i).Value = "CONSULTA";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn07 - 1, i).Value = "";
                    _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn07, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWorkProgress()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void RecordWorkProgress(bool percentUpdate, bool unitsCompletedUpdate)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            _cells.ClearTableRangeColumn(TableName07, ResultColumn07);
            var i = TitleRow07 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wo = new WorkOrder { districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value) };

                    wo.SetWorkOrderDto(WorkOrderActions.GetNewWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value)));
                    wo.pcComplete = percentUpdate ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value) : null;
                    wo.unitsComplete = unitsCompletedUpdate ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value) : null;

                    var reply = WorkOrderActions.RecordWorkProgress(urlService, opSheet, wo.districtCode, wo.GetWorkOrderDto(), wo.pcComplete, wo.unitsComplete);

                    var errorMessage = "";
                    if (percentUpdate && Convert.ToDecimal(wo.pcComplete) != reply.pcComplete)
                        errorMessage += "El valor de Porcentaje Completado no coincide con el ingresado " + wo.pcComplete + " vs " + reply.pcComplete + ".";
                    if (unitsCompletedUpdate && Convert.ToDecimal(wo.unitsComplete) != reply.unitsComplete)
                        errorMessage += "El valor de Unidades Completadas no coincide con el ingresado " + wo.unitsComplete + " vs " + reply.unitsComplete + ".";

                    if (string.IsNullOrWhiteSpace(errorMessage))
                    {
                        _cells.GetCell(4, i).Value = reply.pcComplete;
                        _cells.GetCell(5, i).Value = reply.unitsComplete;
                        _cells.GetCell(ResultColumn07, i).Value = "COMPLETADA";
                        _cells.GetCell(1, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Success;
                    }
                    else
                    {
                        _cells.GetCell(ResultColumn07, i).Value = errorMessage;
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Error;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn07, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateWorkProgress()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn07, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void UpdateWorkProgress()
        {
            if (_progressUpdate)
                RecordWorkProgress(true, false);
            else
                RecordWorkProgress(false, true);
        }

        private void UpdateRequiredProgress()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            _cells.ClearTableRangeColumn(TableName07, ResultColumn07);
            var i = TitleRow07 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var wo = new WorkOrder();
                    //GENERAL
                    wo.districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
                    wo.SetWorkOrderDto(WorkOrderActions.GetNewWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value)));
                    wo.unitOfWork = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    wo.unitsRequired = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);

                    var reply = WorkOrderActions.ModifyWorkOrder(urlService, opSheet, wo);

                    _cells.GetCell(4, i).Value = reply.pcComplete;
                    _cells.GetCell(5, i).Value = reply.unitsComplete;
                    _cells.GetCell(ResultColumn07, i).Value = "COMPLETADA";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn07, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn07, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateWorkProgress()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn07, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void btnCleanWorkOrderSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
            _cells.ClearTableRange(TableNameD01);
            _cells.ClearTableRange(TableNameD02);
            _cells.ClearTableRange(TableNameD03);
            _cells.ClearTableRange(TableNameD04);
        }

        private void btnCleanCloseSheets_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName04);
            _cells.ClearTableRange(TableName05);
        }
        private void btnCleanDuration_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName06);
        }
        private void btnReviewReferenceCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameD04))
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReviewRefCodesList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRefCodesList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnUpdateReferenceCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameD04))
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateReferenceCodes);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewRefCodesList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCleanQualitySheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableNameQ01);
        }
        private void btnReviewQuality_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameQ01))
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewQualityList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewQuality()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReReviewQuality_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameQ01))
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewQualityList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReReviewQuality()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                #region CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 4
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 5
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 6
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 7

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK ORDERS - ELLIPSE 8";
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
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = WoStatusList.GetStatusNames(true);
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.GetCell("A4").AddComment("--ÁREA GERENCIAL/SUPERINTENDENCIA--\n" +
                    "INST: IMIS, MINA\n" +
                    "" + ManagementArea.ManejoDeCarbon.Key + ": " + QuarterMasters.Ferrocarril.Key + ", " + QuarterMasters.PuertoBolivar.Key + ", " + QuarterMasters.PlantasDeCarbon.Key + "\n" +
                    "" + ManagementArea.Mantenimiento.Key + ": MINA\n" +
                    "" + ManagementArea.SoporteOperacion.Key + ": ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A4").Comment.Shape.TextFrame.AutoSize = true;

                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 3, false);
                _cells.GetCell("A5").Value = SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A5"), ValidationSheetName, 2);
                _cells.GetCell("A6").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B6"), statusList, ValidationSheetName, 4);
                _cells.GetRange("A3", "A6").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B6").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.Raised.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 5);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleOptional;
                for (var i = 4; i < ResultColumn01 - 4; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }

                //GENERAL
                _cells.GetCell(1, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow01).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow01).AddComment("Ingrese solo el prefijo si quiere crear una orden con prefijo");
                _cells.GetCell(3, TitleRow01).Value = "WO_STATUS";
                _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(4, TitleRow01 - 2).Value = "GENERAL";
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPTION";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow01).Value = "EQUIPMENT";
                _cells.GetCell(5, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(7, TitleRow01).Value = "MOD_CODE";

                var priorityCodes = MyUtilities.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                var woTypeCodes = MyUtilities.GetCodeList(WoTypeMtType.GetWoTypeList());
                var mtTypeCodes = MyUtilities.GetCodeList(WoTypeMtType.GetMtTypeList());
                var usTypeCodes = MyUtilities.GetCodeList(WorkOrderActions.GetUserStatusCodeList(_eFunctions).ToList());


                _cells.GetCell(8, TitleRow01).Value = "WO_TYPE";
                _cells.GetCell(8, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(8, TitleRow01 + 1), woTypeCodes, ValidationSheetName, 6, false);
                _cells.GetCell(9, TitleRow01).Value = "MT_TYPE";
                _cells.GetCell(9, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(9, TitleRow01 + 1), mtTypeCodes, ValidationSheetName, 7, false);
                _cells.GetCell(10, TitleRow01).Value = "WO_USER_STATUS";
                _cells.SetValidationList(_cells.GetCell(10, TitleRow01 + 1), usTypeCodes, ValidationSheetName, 8, false);
                _cells.GetCell(11, TitleRow01).Value = "RAISED_DATE";
                _cells.GetCell(11, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(12, TitleRow01).Value = "RAISED_TIME";
                _cells.GetCell(12, TitleRow01).AddComment("hhmmss");
                _cells.GetCell(13, TitleRow01).Value = "ORIGINATOR_ID";
                _cells.GetCell(14, TitleRow01).Value = "ORIG_PRIORITY";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(14, TitleRow01 + 1), priorityCodes, ValidationSheetName, 9, false);
                _cells.GetCell(15, TitleRow01).Value = "ORIG_DOC_TYPE";
                _cells.GetCell(16, TitleRow01).Value = "ORIG_DOC_NO";
                _cells.GetCell(17, TitleRow01).Value = "RELATED_WO";
                _cells.GetCell(18, TitleRow01).Value = "WORKREQUEST";
                _cells.GetCell(19, TitleRow01).Value = "STD_JOB";
                _cells.GetCell(20, TitleRow01).Value = "MST";
                _cells.GetCell(20, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(20, TitleRow01 - 1).Value = "N/A";

                _cells.GetRange(4, TitleRow01 - 2, 20, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(4, TitleRow01 - 2, 20, TitleRow01 - 2).Merge();

                //PLANNING
                _cells.GetCell(21, TitleRow01 - 2).Value = "PLANNING";
                _cells.GetCell(21, TitleRow01).Value = "AUTO_REQ";
                _cells.GetCell(21, TitleRow01).AddComment("Y/N");
                _cells.GetCell(22, TitleRow01).Value = "ASSIGN";
                _cells.GetCell(23, TitleRow01).Value = "PLAN_PRIORITY";
                _cells.GetCell(23, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(23, TitleRow01 + 1), ValidationSheetName, 9, false);
                _cells.GetCell(24, TitleRow01).Value = "REQ_START_DATE";
                _cells.GetCell(24, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(25, TitleRow01).Value = "REQ_START_TIME";
                _cells.GetCell(25, TitleRow01).AddComment("hhmmss");
                _cells.GetCell(26, TitleRow01).Value = "REQ_BY_DATE";
                _cells.GetCell(26, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(27, TitleRow01).Value = "REQ_BY_TIME";
                _cells.GetCell(27, TitleRow01).AddComment("hhmmss");
                _cells.GetCell(28, TitleRow01).Value = "PLAN_STR_DATE";
                _cells.GetCell(28, TitleRow01).AddComment("yyyyMMdd - Las fechas de plan solo se modificarán si el usuario tiene permisos de planeación/programación");
                _cells.GetCell(29, TitleRow01).Value = "PLAN_STR_TIME";
                _cells.GetCell(29, TitleRow01).AddComment("hhmmss");
                _cells.GetCell(30, TitleRow01).Value = "PLAN_FIN_DATE";
                _cells.GetCell(30, TitleRow01).AddComment("yyyyMMdd - El comportamiento de este campo depende de la tarea de la orden");
                _cells.GetCell(31, TitleRow01).Value = "PLAN_FIN_TIME";
                _cells.GetCell(31, TitleRow01).AddComment("hhmmss");
                _cells.GetCell(32, TitleRow01).Value = "UNIT_OF_WORK";
                _cells.GetCell(33, TitleRow01).Value = "UNITS_REQUIRED";
                _cells.GetCell(34, TitleRow01).Value = "PC/UNITS COMP";
                _cells.GetCell(34, TitleRow01 - 1).Value = "N/A";
                _cells.GetCell(34, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetRange(21, TitleRow01 - 2, 34, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(21, TitleRow01 - 2, 34, TitleRow01 - 2).Merge();

                //COST
                _cells.GetCell(35, TitleRow01 - 2).Value = "COST";
                _cells.GetCell(35, TitleRow01).Value = "ACCOUNT_CODE";
                _cells.GetCell(36, TitleRow01).Value = "PROJECT_NO";
                _cells.GetCell(37, TitleRow01).Value = "PARENT_WO";
                _cells.GetRange(35, TitleRow01 - 2, 37, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(35, TitleRow01 - 2, 37, TitleRow01 - 2).Merge();

                //JOB_CODES
                _cells.GetCell(38, TitleRow01 - 2).Value = "JOB CODES/FALLAS";
                _cells.GetCell(38, TitleRow01 - 2).AddComment("Debe seleccionar por lo menos un Job Code para las órdenes correctivas/reparación");
                _cells.GetCell(38, TitleRow01).Value = "JOBCODE_01";
                _cells.GetCell(39, TitleRow01).Value = "JOBCODE_02";
                _cells.GetCell(40, TitleRow01).Value = "JOBCODE_03";
                _cells.GetCell(41, TitleRow01).Value = "JOBCODE_04";
                _cells.GetCell(42, TitleRow01).Value = "JOBCODE_05";
                _cells.GetCell(43, TitleRow01).Value = "JOBCODE_06";
                _cells.GetCell(44, TitleRow01).Value = "JOBCODE_07";
                _cells.GetCell(45, TitleRow01).Value = "JOBCODE_08";
                _cells.GetCell(46, TitleRow01).Value = "JOBCODE_09";
                _cells.GetCell(47, TitleRow01).Value = "JOBCODE_10";
                _cells.GetCell(48, TitleRow01).Value = "LOCATION FR";
                _cells.GetCell(49, TitleRow01).Value = "PART FAILURE";
                _cells.GetRange(38, TitleRow01 - 2, 49, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(38, TitleRow01 - 2, 49, TitleRow01 - 2).Merge();
                //COMPLETION INFO
                _cells.GetCell(50, TitleRow01 - 2).Value = "COMPL.INFO";
                _cells.GetCell(50, TitleRow01).Value = "COMPL_COD";
                _cells.GetCell(50, TitleRow01).AddComment("Código de cierre de la orden");
                _cells.GetCell(51, TitleRow01).Value = "COMP_COMM";
                _cells.GetCell(51, TitleRow01).AddComment("Indica si una orden tiene comentarios de cierre");
                _cells.GetCell(52, TitleRow01).Value = "CLOSED DATE";
                _cells.GetCell(53, TitleRow01).Value = "COMPL_BY";
                _cells.GetRange(50, TitleRow01 - 2, 53, TitleRow01 - 2).Style = StyleConstants.Option;
                _cells.GetRange(50, TitleRow01, 53, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetRange(50, TitleRow01 - 2, 53, TitleRow01 - 1).Merge();

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 2
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WO TASKS - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetRange(1, TitleRow02, ResultColumn02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, TitleRow02 - 1).Value = "WORK ORDER";
                _cells.GetRange(1, TitleRow02 - 1, 5, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, TitleRow02 - 1, 5, TitleRow02 - 1).Merge();

                _cells.GetCell(1, TitleRow02).Value = "DISTRICT";
                _cells.GetCell(2, TitleRow02).Value = "WORK_GROUP";
                _cells.GetCell(3, TitleRow02).Value = "WORK_ORDER";
                _cells.GetCell(4, TitleRow02).Value = "WO_DESC";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleInformation;
                //ACTION
                _cells.GetCell(5, TitleRow02).Value = "ACTION";
                _cells.GetCell(5, TitleRow02).Style = StyleConstants.TitleAction;
                _cells.GetCell(5, TitleRow02).AddComment("C: Crear \nM: Modificar \nD: Eliminar");
                _cells.SetValidationList(_cells.GetCell(5, TitleRow02 + 1), new List<string> { "C", "M", "D" });
                //GENERAL
                _cells.GetCell(6, TitleRow02 - 1).Value = "GENERAL";
                _cells.GetRange(6, TitleRow02 - 1, 11, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(6, TitleRow02 - 1, 11, TitleRow02 - 1).Merge();

                _cells.GetCell(6, TitleRow02).Value = "TASK_NO";
                _cells.GetCell(7, TitleRow02).Value = "WO_TASK_DESC";
                _cells.GetCell(8, TitleRow02).Value = "JOB_DESC_CODE";
                _cells.GetCell(9, TitleRow02).Value = "SAFETY_INST";
                _cells.GetCell(10, TitleRow02).Value = "COMPL_INST";
                _cells.GetCell(11, TitleRow02).Value = "COMPL_TEXT_CODE";

                //PLANNING
                _cells.GetCell(12, TitleRow02 - 1).Value = "PLANNING";
                _cells.GetRange(12, TitleRow02 - 1, 14, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(12, TitleRow02 - 1, 14, TitleRow02 - 1).Merge();

                _cells.GetCell(12, TitleRow02).Value = "ASSIGN_PERSON";
                _cells.GetCell(13, TitleRow02).Value = "EST_MACH_HRS";
                _cells.GetCell(14, TitleRow02).Value = "PLAN START DATE";
                _cells.GetRange(12, TitleRow02, 14, TitleRow02).Style = StyleConstants.TitleOptional;

                //RECURSOS
                _cells.GetCell(15, TitleRow02 - 1).Value = "RECURSOS";
                _cells.GetRange(15, TitleRow02 - 1, 17, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(15, TitleRow02 - 1, 17, TitleRow02 - 1).Merge();

                _cells.GetCell(15, TitleRow02).Value = "EST_DUR_HRS";
                _cells.GetCell(15, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(16, TitleRow02).Value = "LABOR";
                _cells.GetCell(17, TitleRow02).Value = "MATERIAL";
                _cells.GetRange(15, TitleRow02, 17, TitleRow02).Style = StyleConstants.TitleInformation;

                //APL
                _cells.GetCell(18, TitleRow02 - 1).Value = "APL";
                _cells.GetRange(18, TitleRow02 - 1, 22, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(18, TitleRow02 - 1, 22, TitleRow02 - 1).Merge();

                _cells.GetCell(18, TitleRow02).Value = "EQUIP_GRP_ID";
                _cells.GetCell(19, TitleRow02).Value = "APL_TYPE";
                _cells.GetCell(20, TitleRow02).Value = "COMP_CODE";
                _cells.GetCell(21, TitleRow02).Value = "COMP_MOD_CODE";
                _cells.GetCell(22, TitleRow02).Value = "APL_SEQ_NO";

                _cells.GetRange(18, TitleRow02, 22, TitleRow02).Style = StyleConstants.TitleOptional;


                _cells.GetCell(23, TitleRow02).Value = "DESCRIPCION EXTENDIDA";
                _cells.GetCell(23, TitleRow02).Style = StyleConstants.TitleOptional;
                //RESULTADO
                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 3
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WO TASK REQUIREMENTS - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetRange(1, TitleRow03, ResultColumn03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, TitleRow03 - 1).Value = "WO / TASK";
                _cells.GetRange(1, TitleRow03 - 1, 6, TitleRow03 - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, TitleRow03 - 1, 6, TitleRow03 - 1).Merge();

                _cells.GetCell(1, TitleRow03).Value = "DISTRICT";       //_cells.GetCell(1, i).Value = req.DistrictCode; 
                _cells.GetCell(2, TitleRow03).Value = "WORK_GROUP";     //_cells.GetCell(2, i).Value = req.WorkGroup;    
                _cells.GetCell(3, TitleRow03).Value = "WO_NO";          //_cells.GetCell(3, i).Value = req.WorkOrder;    
                _cells.GetCell(4, TitleRow03).Value = "TASK_NO";        //_cells.GetCell(4, i).Value = req.WoTaskNo;     
                _cells.GetCell(5, TitleRow03).Value = "WO_TASK_DESC";   //_cells.GetCell(5, i).Value = req.WoTaskDesc;    

                //ACTION
                _cells.GetCell(6, TitleRow03).Value = "ACTION";
                _cells.GetCell(6, TitleRow03).Style = StyleConstants.TitleAction;
                _cells.GetCell(6, TitleRow03).AddComment("C: Crear Requerimiento \nM: Modificar Requerimiento \nD: Eliminar Requerimiento");
                _cells.SetValidationList(_cells.GetCell(6, TitleRow03 + 1), new List<string> { "C", "M", "D" });
                //GENERAL
                _cells.GetCell(7, TitleRow03 - 1).Value = "GENERAL";
                _cells.GetRange(7, TitleRow03 - 1, 14, TitleRow03 - 1).Style = StyleConstants.Option;
                _cells.GetRange(7, TitleRow03 - 1, 14, TitleRow03 - 1).Merge();

                _cells.GetCell(7, TitleRow03).Value = "REQ_TYPE";       //_cells.GetCell(7, i).Value = "" + req.ReqType;
                _cells.GetCell(7, TitleRow03).AddComment("LAB: LABOR\nMAT: MATERIAL");
                _cells.SetValidationList(_cells.GetCell(7, TitleRow03 + 1), new List<string> { "LAB", "MAT" });


                _cells.GetCell(8, TitleRow03).Value = "SEQ_NO";         //_cells.GetCell(8, i).Value = req.SeqNo;    
                _cells.GetCell(9, TitleRow03).Value = "REQ_CODE";       //_cells.GetCell(9, i).Value = req.ReqCode;  
                _cells.GetCell(10, TitleRow03).Value = "DESCRIPTION";   //_cells.GetCell(10, i).Value = req.ReqDesc; 
                _cells.GetCell(11, TitleRow03).Value = "UOM";           //_cells.GetCell(11, i).Value = req.UoM;  
                _cells.GetCell(12, TitleRow03).Value = "QTY REQ";       //_cells.GetCell(11, i).Value = req.QtyReq;  
                _cells.GetCell(13, TitleRow03).Value = "QTY ISS";       //_cells.GetCell(12, i).Value = req.QtyIss;  
                _cells.GetCell(14, TitleRow03).Value = "HRS_REQ";       //_cells.GetCell(13, i).Value = req.HrsReq;  
                _cells.GetCell(15, TitleRow03).Value = "HRS_REAL";      //_cells.GetCell(14, i).Value = req.HrsReal; 


                _cells.GetCell(8, TitleRow03).AddComment("Aplica solo para Creación y Modificación de Requerimientos");
                _cells.GetCell(9, TitleRow03).AddComment("Recurso: Class+Code (Ver hoja de recursos) \nMaterial: StockCode");
                _cells.GetCell(12, TitleRow03).AddComment("Horas requeridas del recurso. (Solo aplica para labor)");
                _cells.GetCell(13, TitleRow03).AddComment("Horas Reales del recurso. (Solo aplica para labor)");
                _cells.GetCell(14, TitleRow03).AddComment("Unidad de Medida. (Solo aplica para Equipos)");


                //RESULTADO
                _cells.GetCell(ResultColumn03, TitleRow03).Value = "RESULTADO";
                _cells.GetCell(ResultColumn03, TitleRow03).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow03 + 1, ResultColumn03 - 2, TitleRow03 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumn03, TitleRow03 + 1), TableName03);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 4 - CLOSE WO
                _excelApp.ActiveWorkbook.Sheets[4].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName04;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "CLOSE WORK ORDERS - ELLIPSE 8";
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
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);

                //GENERAL
                _cells.GetRange(1, TitleRow04, ResultColumn04 - 1, TitleRow04).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow04).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow04).Value = "CLOSED_DATE";
                _cells.GetCell(2, TitleRow04).AddComment("yyyyMMdd");
                _cells.GetCell(3, TitleRow04).Value = "CLOSED_TIME";
                _cells.GetCell(3, TitleRow04).AddComment("hhmmss");
                _cells.GetCell(3, TitleRow04).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, TitleRow04).Value = "COMPLETED_BY";
                _cells.GetCell(5, TitleRow04).Value = "COMPLETED_CODE";
                var completeCodeList = _eFunctions.GetItemCodes("SC").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(5, TitleRow04 + 1), completeCodeList, ValidationSheetName, 10, false);
                _cells.GetCell(6, TitleRow04).Value = "OUT_SERV_DATE";
                _cells.GetCell(6, TitleRow04).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, TitleRow04).Value = "COMENTARIO";
                _cells.GetCell(7, TitleRow04).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, TitleRow04).AddComment("Adiciona el siguiente texto al campo de comentario (no elimina el comentario existente)");

                _cells.GetCell(ResultColumn04, TitleRow04).Value = "RESULTADO";
                _cells.GetCell(ResultColumn04, TitleRow04).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow04 + 1, ResultColumn04, TitleRow04 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow04, ResultColumn04, TitleRow04 + 1), TableName04);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 5 - CLOSE COMMENTS
                _excelApp.ActiveWorkbook.Sheets[5].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName05;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "CLOSE WORK ORDERS - ELLIPSE 8";
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
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);

                _cells.GetRange(1, TitleRow05, ResultColumn05 - 1, TitleRow05).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow05).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow05).Value = "COMENTARIO";
                _cells.GetCell(2, TitleRow05).Style = StyleConstants.TitleOptional;

                _cells.GetCell(ResultColumn05, TitleRow05).Value = "RESULTADO";
                _cells.GetCell(ResultColumn05, TitleRow05).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow05 + 1, ResultColumn05, TitleRow05 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow05, ResultColumn05, TitleRow05 + 1), TableName05);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 6 - DURATION
                _excelApp.ActiveWorkbook.Sheets[6].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName06;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "WORK ORDERS DURATIONS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                //GENERAL
                _cells.GetRange(1, TitleRow06, ResultColumn04 - 1, TitleRow06).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, TitleRow06).Value = "DISTRICT_CODE";
                _cells.GetCell(2, TitleRow06).Value = "WORK_ORDER";
                _cells.GetCell(3, TitleRow06).Value = "DURATION_DATE";
                _cells.GetCell(3, TitleRow06).AddComment("yyyyMMdd");
                _cells.GetCell(4, TitleRow06).Value = "DURATION_CODE";
                var durationCodeList = _eFunctions.GetItemCodes("JI").Select(item => item.code + " - " + item.description).ToList();
                _cells.SetValidationList(_cells.GetCell(4, TitleRow06 + 1), durationCodeList, ValidationSheetName, 11, false);
                _cells.GetCell(5, TitleRow06).Value = "START_HOUR";
                _cells.GetCell(5, TitleRow06).AddComment("hhmmss");
                _cells.GetCell(6, TitleRow06).Value = "FINAL_HOUR";
                _cells.GetCell(6, TitleRow06).AddComment("hhmmss");
                _cells.GetCell(7, TitleRow06).Value = "ACTION";
                _cells.GetCell(7, TitleRow06).Style = StyleConstants.TitleAction;
                _cells.GetCell(7, TitleRow06).AddComment("Crear, Eliminar");
                var actionsList = new List<string> { "Crear", "Eliminar" };
                _cells.SetValidationList(_cells.GetCell(7, TitleRow06 + 1), actionsList, ValidationSheetName, 12, false);

                _cells.GetCell(ResultColumn06, TitleRow06).Value = "RESULTADO";
                _cells.GetCell(ResultColumn06, TitleRow06).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow06 + 1, ResultColumn06, TitleRow06 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow06, ResultColumn06, TitleRow06 + 1), TableName06);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 7 - PROGRESS WO
                _excelApp.ActiveWorkbook.Sheets[7].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName07;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "WORK ORDERS PROGRESS - ELLIPSE 8";
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
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);

                //GENERAL
                _cells.GetRange(1, TitleRow07, ResultColumn07 - 1, TitleRow07).Style = StyleConstants.TitleOptional;
                _cells.GetCell(1, TitleRow07).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow07).Value = "UNITS OF WORK";
                _cells.GetCell(3, TitleRow07).Value = "UNITS REQUIRED";
                _cells.GetCell(4, TitleRow07).Value = "PERCENT COMPLETED";
                _cells.GetCell(5, TitleRow07).Value = "UNITS COMPLETED";

                _cells.GetCell(ResultColumn07, TitleRow07).Value = "RESULTADO";
                _cells.GetCell(ResultColumn07, TitleRow07).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow07 + 1, ResultColumn07, TitleRow07 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow07, ResultColumn07, TitleRow07 + 1), TableName07);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void FormatDetailed()
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

                _cells.SetCursorWait();
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 4
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameD01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK ORDERS - ELLIPSE 8";
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
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = WoStatusList.GetStatusNames(true);
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.GetCell("A4").AddComment("--ÁREA/SUPERINTENDENCIA--\n" +
                    "INST: IMIS, MINA\n" +
                    "MDC: FFCC, PBV, PTAS\n" +
                    "MNTTO: MINA\n" +
                    "SOP: ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A4").Comment.Shape.TextFrame.AutoSize = true;

                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 3, false);
                _cells.GetCell("A5").Value = SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A5"), ValidationSheetName, 2);
                _cells.GetCell("A6").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B6"), statusList, ValidationSheetName, 4);
                _cells.GetRange("A3", "A6").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B6").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.Raised.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 5);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetRange(1, TitleRowD01, ResultColumnD01, TitleRowD01).Style = StyleConstants.TitleOptional;
                for (var i = 4; i < ResultColumnD01 - 6; i++)
                {
                    _cells.GetCell(i, TitleRowD01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRowD01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRowD01 - 1).Value = "true";
                }
                for (var i = ResultColumnD01 - 2; i < ResultColumnD01; i++)
                {
                    _cells.GetCell(i, TitleRowD01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRowD01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRowD01 - 1).Value = "true";
                }

                //GENERAL
                _cells.GetCell(1, TitleRowD01).Value = "WORK_GROUP";
                _cells.GetCell(1, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRowD01).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRowD01).AddComment("Ingrese solo el prefijo si quiere crear una orden con prefijo");
                _cells.GetCell(3, TitleRowD01).Value = "WO_STATUS";
                _cells.GetCell(3, TitleRowD01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(4, TitleRowD01 - 2).Value = "GENERAL";
                _cells.GetCell(4, TitleRowD01).Value = "DESCRIPTION";
                _cells.GetCell(4, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRowD01).Value = "EQUIPMENT";
                _cells.GetCell(5, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, TitleRowD01).Value = "COMP_CODE";
                _cells.GetCell(7, TitleRowD01).Value = "MOD_CODE";

                var priorityCodes = MyUtilities.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                var woTypeCodes = MyUtilities.GetCodeList(WoTypeMtType.GetWoTypeList());
                var mtTypeCodes = MyUtilities.GetCodeList(WoTypeMtType.GetMtTypeList());
                var usTypeCodes = MyUtilities.GetCodeList(WorkOrderActions.GetUserStatusCodeList(_eFunctions).ToList());
                var contactMethod = MyUtilities.GetCodeList(_eFunctions.GetItemCodes("MTCO"));


                _cells.GetCell(8, TitleRowD01).Value = "WO_TYPE";
                _cells.GetCell(8, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(8, TitleRowD01 + 1), woTypeCodes, ValidationSheetName, 6, false);
                _cells.GetCell(9, TitleRowD01).Value = "MT_TYPE";
                _cells.GetCell(9, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(9, TitleRowD01 + 1), mtTypeCodes, ValidationSheetName, 7, false);
                _cells.GetCell(10, TitleRowD01).Value = "WO_USER_STATUS";
                _cells.SetValidationList(_cells.GetCell(10, TitleRowD01 + 1), usTypeCodes, ValidationSheetName, 8, false);
                _cells.GetCell(11, TitleRowD01).Value = "RAISED_DATE";
                _cells.GetCell(11, TitleRowD01).AddComment("yyyyMMdd");
                _cells.GetCell(12, TitleRowD01).Value = "RAISED_TIME";
                _cells.GetCell(12, TitleRowD01).AddComment("hhmmss");
                _cells.GetCell(13, TitleRowD01).Value = "ORIGINATOR_ID";
                _cells.GetCell(14, TitleRowD01).Value = "ORIG_PRIORITY";
                _cells.GetCell(14, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(14, TitleRowD01 + 1), priorityCodes, ValidationSheetName, 9, false);
                _cells.GetCell(15, TitleRowD01).Value = "ORIG_DOC_TYPE";
                _cells.GetCell(16, TitleRowD01).Value = "ORIG_DOC_NO";
                _cells.GetCell(17, TitleRowD01).Value = "RELATED_WO";
                _cells.GetCell(18, TitleRowD01).Value = "WORKREQUEST";
                _cells.GetCell(19, TitleRowD01).Value = "STD_JOB";
                _cells.GetCell(20, TitleRowD01).Value = "MST";
                _cells.GetCell(20, TitleRowD01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(20, TitleRowD01 - 1).Value = "N/A";

                _cells.GetRange(4, TitleRowD01 - 2, 20, TitleRowD01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(4, TitleRowD01 - 2, 20, TitleRowD01 - 2).Merge();

                //PLANNING
                _cells.GetCell(21, TitleRowD01 - 2).Value = "PLANNING";
                _cells.GetCell(21, TitleRowD01).Value = "AUTO_REQ";
                _cells.GetCell(21, TitleRowD01).AddComment("Y/N");
                _cells.GetCell(22, TitleRowD01).Value = "ASSIGN";
                _cells.GetCell(23, TitleRowD01).Value = "PLAN_PRIORITY";
                _cells.GetCell(23, TitleRowD01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(23, TitleRowD01 + 1), ValidationSheetName, 9, false);
                _cells.GetCell(24, TitleRowD01).Value = "REQ_START_DATE";
                _cells.GetCell(24, TitleRowD01).AddComment("yyyyMMdd");
                _cells.GetCell(25, TitleRowD01).Value = "REQ_START_TIME";
                _cells.GetCell(25, TitleRowD01).AddComment("hhmmss");
                _cells.GetCell(26, TitleRowD01).Value = "REQ_BY_DATE";
                _cells.GetCell(26, TitleRowD01).AddComment("yyyyMMdd");
                _cells.GetCell(27, TitleRowD01).Value = "REQ_BY_TIME";
                _cells.GetCell(27, TitleRowD01).AddComment("hhmmss");
                _cells.GetCell(28, TitleRowD01).Value = "PLAN_STR_DATE";
                _cells.GetCell(28, TitleRowD01).AddComment("yyyyMMdd - Las fechas de plan solo se modificarán si el usuario tiene permisos de planeación/programación");
                _cells.GetCell(29, TitleRowD01).Value = "PLAN_STR_TIME";
                _cells.GetCell(29, TitleRowD01).AddComment("hhmmss");
                _cells.GetCell(30, TitleRowD01).Value = "PLAN_FIN_DATE";
                _cells.GetCell(30, TitleRowD01).AddComment("yyyyMMdd - El comportamiento de este campo depende de la tarea de la orden");
                _cells.GetCell(31, TitleRowD01).Value = "PLAN_FIN_TIME";
                _cells.GetCell(31, TitleRowD01).AddComment("hhmmss");
                _cells.GetCell(32, TitleRowD01).Value = "UNIT_OF_WORK";
                _cells.GetCell(33, TitleRowD01).Value = "UNITS_REQUIRED";
                _cells.GetCell(34, TitleRowD01).Value = "PC/UNITS COMP";
                _cells.GetCell(34, TitleRowD01 - 1).Value = "N/A";
                _cells.GetCell(34, TitleRowD01).Style = StyleConstants.TitleInformation;
                _cells.GetRange(21, TitleRowD01 - 2, 34, TitleRowD01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(21, TitleRowD01 - 2, 34, TitleRowD01 - 2).Merge();

                //COST
                _cells.GetCell(35, TitleRowD01 - 2).Value = "COST";
                _cells.GetCell(35, TitleRowD01).Value = "ACCOUNT_CODE";
                _cells.GetCell(36, TitleRowD01).Value = "PROJECT_NO";
                _cells.GetCell(37, TitleRowD01).Value = "PARENT_WO";
                _cells.GetRange(35, TitleRowD01 - 2, 37, TitleRowD01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(35, TitleRowD01 - 2, 37, TitleRowD01 - 2).Merge();

                //JOB_CODES
                _cells.GetCell(38, TitleRowD01 - 2).Value = "JOB CODES/FALLAS";
                _cells.GetCell(38, TitleRowD01 - 2).AddComment("Debe seleccionar por lo menos un Job Code para las órdenes correctivas/reparación");
                _cells.GetCell(38, TitleRowD01).Value = "JOBCODE_D01";
                _cells.GetCell(39, TitleRowD01).Value = "JOBCODE_02";
                _cells.GetCell(40, TitleRowD01).Value = "JOBCODE_03";
                _cells.GetCell(41, TitleRowD01).Value = "JOBCODE_04";
                _cells.GetCell(42, TitleRowD01).Value = "JOBCODE_05";
                _cells.GetCell(43, TitleRowD01).Value = "JOBCODE_06";
                _cells.GetCell(44, TitleRowD01).Value = "JOBCODE_07";
                _cells.GetCell(45, TitleRowD01).Value = "JOBCODE_08";
                _cells.GetCell(46, TitleRowD01).Value = "JOBCODE_09";
                _cells.GetCell(47, TitleRowD01).Value = "JOBCODE_10";
                _cells.GetCell(48, TitleRowD01).Value = "LOCATION FR";
                _cells.GetCell(49, TitleRowD01).Value = "PART FAILURE";
                _cells.GetRange(38, TitleRowD01 - 2, 49, TitleRowD01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(38, TitleRowD01 - 2, 49, TitleRowD01 - 2).Merge();
                //COMPLETION INFO
                _cells.GetCell(50, TitleRowD01 - 2).Value = "COMPL.INFO";
                _cells.GetCell(50, TitleRowD01).Value = "COMPL_COD";
                _cells.GetCell(50, TitleRowD01).AddComment("Código de cierre de la orden");
                _cells.GetCell(51, TitleRowD01).Value = "COMP_COMM";
                _cells.GetCell(51, TitleRowD01).AddComment("Indica si una orden tiene comentarios de cierre");
                _cells.GetCell(52, TitleRowD01).Value = "CLOSED DATE";
                _cells.GetCell(53, TitleRowD01).Value = "COMPL_BY";
                _cells.GetRange(50, TitleRowD01 - 2, 53, TitleRowD01 - 2).Style = StyleConstants.Option;
                _cells.GetRange(50, TitleRowD01, 53, TitleRowD01).Style = StyleConstants.TitleInformation;
                _cells.GetRange(50, TitleRowD01 - 2, 53, TitleRowD01 - 1).Merge();
                //EXTRA
                _cells.GetCell(54, TitleRowD01 - 2).Value = "EXTRA";
                _cells.GetCell(54, TitleRowD01).Value = "DESC EXT HEADER";
                _cells.GetCell(55, TitleRowD01).Value = "DESC EXT BODY";
                _cells.GetRange(54, TitleRowD01 + 1, 55, TitleRowD01 + 1).WrapText = false;
                _cells.GetRange(54, TitleRowD01 - 2, 55, TitleRowD01 - 2).Style = StyleConstants.Select;
                _cells.GetRange(54, TitleRowD01 - 2, 55, TitleRowD01 - 2).Merge();
                _cells.GetCell(ResultColumnD01, TitleRowD01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnD01, TitleRowD01).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRowD01 + 1, ResultColumnD01, TitleRowD01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowD01, ResultColumnD01, TitleRowD01 + 1), TableNameD01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 2 - WO TASKS

                //CONSTRUYO LA HOJA 3 - WO TASKS REQUIREMENTS

                //CONSTRUYO LA HOJA 4 - WO TASKS REFERENCE CODES
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameD04;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "WORK ORDER REFERENCE CODES - ELLIPSE 8";
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

                _cells.GetRange(1, TitleRowD04, 3, TitleRowD04).Style = StyleConstants.TitleRequired;
                _cells.GetRange(3, TitleRowD04, ResultColumnD04 - 1, TitleRowD04).Style = StyleConstants.TitleOptional;
                for (var i = 3; i < ResultColumnD04; i++)
                {
                    _cells.GetCell(i, TitleRowD04 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRowD04 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRowD04 - 1).Value = "true";
                }

                _cells.GetCell(1, TitleRowD04).Value = "District";
                _cells.GetCell(2, TitleRowD04).Value = "Work Order";
                _cells.GetCell(3, TitleRowD04).Value = "Work Request";
                _cells.GetCell(4, TitleRowD04).Value = "Comentarios Duraciones";
                _cells.GetCell(5, TitleRowD04).Value = "Com.Dur. Text";
                _cells.GetCell(6, TitleRowD04).Value = "EmpleadoId";
                _cells.GetCell(7, TitleRowD04).Value = "Nro. Componente";
                _cells.GetCell(8, TitleRowD04).Value = "P1. Eq.Liv-Med";
                _cells.GetCell(9, TitleRowD04).Value = "P2. Eq.Movil-Minero";
                _cells.GetCell(10, TitleRowD04).Value = "P3. Manejo Sust.Peligrosa";
                _cells.GetCell(11, TitleRowD04).Value = "P4. Guardas Equipo";
                _cells.GetCell(12, TitleRowD04).Value = "P5. Aislamiento";
                _cells.GetCell(13, TitleRowD04).Value = "P6. Trabajos Altura";
                _cells.GetCell(14, TitleRowD04).Value = "P7. Manejo Cargas";
                _cells.GetCell(15, TitleRowD04).Value = "Proyecto ICN";
                _cells.GetCell(16, TitleRowD04).Value = "Reembolsable";
                _cells.GetCell(17, TitleRowD04).Value = "Fecha No Conforme";
                _cells.GetCell(18, TitleRowD04).Value = "Fecha NC Text";
                _cells.GetCell(19, TitleRowD04).Value = "No Conforme?";
                _cells.GetCell(20, TitleRowD04).Value = "Fecha Ejecución";
                _cells.GetCell(21, TitleRowD04).Value = "Hora Ingreso";
                _cells.GetCell(22, TitleRowD04).Value = "Hora Salida";
                _cells.GetCell(23, TitleRowD04).Value = "Nombre Buque";
                _cells.GetCell(24, TitleRowD04).Value = "Calif. Encuesta";
                _cells.GetCell(25, TitleRowD04).Value = "Tarea Crítica?";
                _cells.GetCell(26, TitleRowD04).Value = "Garantía";
                _cells.GetCell(27, TitleRowD04).Value = "Garantía Text";
                _cells.GetCell(28, TitleRowD04).Value = "Cód. Certificación";
                _cells.GetCell(29, TitleRowD04).Value = "Fecha Entrega";
                _cells.GetCell(30, TitleRowD04).Value = "Relacionar EV";
                _cells.GetCell(31, TitleRowD04).Value = "Departamento";
                _cells.GetCell(32, TitleRowD04).Value = "Localización";
                _cells.GetCell(33, TitleRowD04).Value = "Metodo de Contacto";
                _cells.SetValidationList(_cells.GetCell(33, TitleRowD04 + 1), contactMethod);
                _cells.GetCell(34, TitleRowD04).Value = "Detalle de Contacto";


                _cells.GetCell(ResultColumnD04, TitleRowD04).Value = "RESULTADO";
                _cells.GetCell(ResultColumnD04, TitleRowD04).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRowD04 + 1, ResultColumnD04, TitleRowD04 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRowD04, ResultColumnD04, TitleRowD04 + 1), TableNameD04);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatDetailed()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void ReviewWoDetailedList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            //Obtengo los valores de las opciones de búsqueda
            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A5").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B6").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            var listwo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
            var i = TitleRow01 + 1;
            foreach (var wo in listwo)
            {
                try
                {
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "" + wo.workGroup;
                    _cells.GetCell(2, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(3, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(4, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(5, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(6, i).Value = "'" + wo.compCode;
                    _cells.GetCell(7, i).Value = "'" + wo.compModCode;
                    _cells.GetCell(8, i).Value = "" + wo.workOrderType;
                    _cells.GetCell(9, i).Value = "" + wo.maintenanceType;
                    _cells.GetCell(10, i).Value = "" + wo.workOrderStatusU;
                    //DETAILS
                    _cells.GetCell(11, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(12, i).Value = "'" + wo.raisedTime;
                    _cells.GetCell(13, i).Value = "" + wo.originatorId;
                    _cells.GetCell(14, i).Value = "" + wo.origPriority;
                    _cells.GetCell(15, i).Value = "" + wo.origDocType;
                    _cells.GetCell(16, i).Value = "'" + wo.origDocNo;
                    _cells.GetCell(17, i).Value = "'" + wo.GetRelatedWoDto().prefix + wo.GetRelatedWoDto().no;
                    _cells.GetCell(18, i).Value = "'" + wo.requestId;
                    _cells.GetCell(19, i).Value = "'" + wo.stdJobNo;
                    _cells.GetCell(20, i).Value = "'" + wo.maintSchTask;
                    //PLANNING
                    _cells.GetCell(21, i).Value = "" + wo.autoRequisitionInd;
                    _cells.GetCell(22, i).Value = "" + wo.assignPerson;
                    _cells.GetCell(23, i).Value = "" + wo.planPriority;
                    _cells.GetCell(24, i).Value = "'" + wo.requisitionStartDate;
                    _cells.GetCell(25, i).Value = "'" + wo.requisitionStartTime;
                    _cells.GetCell(26, i).Value = "'" + wo.requiredByDate;
                    _cells.GetCell(27, i).Value = "'" + wo.requiredByTime;
                    _cells.GetCell(28, i).Value = "'" + wo.planStrDate;
                    _cells.GetCell(29, i).Value = "'" + wo.planStrTime;
                    _cells.GetCell(30, i).Value = "'" + wo.planFinDate;
                    _cells.GetCell(31, i).Value = "'" + wo.planFinTime;

                    _cells.GetCell(32, i).Value = "" + wo.unitOfWork;
                    _cells.GetCell(33, i).Value = "" + wo.unitsRequired;
                    _cells.GetCell(34, i).Value = "" + wo.pcComplete + "%/" + wo.unitsComplete;
                    //COST
                    _cells.GetCell(35, i).Value = "" + wo.accountCode;
                    _cells.GetCell(36, i).Value = "" + wo.projectNo;
                    _cells.GetCell(37, i).Value = "'" + wo.parentWo;
                    //JOB_CODES
                    _cells.GetCell(38, i).Value2 = "'" + wo.jobCode1;
                    _cells.GetCell(39, i).Value2 = "'" + wo.jobCode2;
                    _cells.GetCell(40, i).Value2 = "'" + wo.jobCode3;
                    _cells.GetCell(41, i).Value2 = "'" + wo.jobCode4;
                    _cells.GetCell(42, i).Value2 = "'" + wo.jobCode5;
                    _cells.GetCell(43, i).Value = "'" + wo.jobCode6;
                    _cells.GetCell(44, i).Value = "'" + wo.jobCode7;
                    _cells.GetCell(45, i).Value = "'" + wo.jobCode8;
                    _cells.GetCell(46, i).Value = "'" + wo.jobCode9;
                    _cells.GetCell(47, i).Value = "'" + wo.jobCode10;
                    _cells.GetCell(48, i).Value = "'" + wo.locationFr;
                    _cells.GetCell(49, i).Value = "'" + wo.failurePart;
                    _cells.GetCell(50, i).Value = "'" + wo.completedCode;
                    _cells.GetCell(51, i).Value = "" + wo.completeTextFlag;
                    _cells.GetCell(52, i).Value = "" + wo.closeCommitDate;
                    _cells.GetCell(53, i).Value = "" + wo.completedBy;
                    _cells.GetCell(54, i).Value = "" + wo.GetExtendedDescription(urlService, opContext).Header;
                    _cells.GetCell(55, i).Value = "" + wo.GetExtendedDescription(urlService, opContext).Body;
                    _cells.GetRange(54, i, 55, i).WrapText = false;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }
            //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit(); //Se desactiva por los comentarios extendidos
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ReReviewWoDetailedList()
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNameD01, ResultColumnD01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opContext = new WorkOrderService.OperationContext
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

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                    district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;
                    var woNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, woNo);

                    if (wo == null || wo.GetWorkOrderDto().no == null)
                        throw new Exception("WORK ORDER NO ENCONTRADA");
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumnD01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "" + wo.workGroup;
                    _cells.GetCell(2, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(3, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(4, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(5, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(6, i).Value = "'" + wo.compCode;
                    _cells.GetCell(7, i).Value = "'" + wo.compModCode;
                    _cells.GetCell(8, i).Value = "" + wo.workOrderType;
                    _cells.GetCell(9, i).Value = "" + wo.maintenanceType;
                    _cells.GetCell(10, i).Value = "" + wo.workOrderStatusU;
                    //DETAILS
                    _cells.GetCell(11, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(12, i).Value = "'" + wo.raisedTime;
                    _cells.GetCell(13, i).Value = "" + wo.originatorId;
                    _cells.GetCell(14, i).Value = "" + wo.origPriority;
                    _cells.GetCell(15, i).Value = "" + wo.origDocType;
                    _cells.GetCell(16, i).Value = "'" + wo.origDocNo;
                    _cells.GetCell(17, i).Value = "'" + wo.GetRelatedWoDto().prefix + wo.GetRelatedWoDto().no;
                    _cells.GetCell(18, i).Value = "'" + wo.requestId;
                    _cells.GetCell(19, i).Value = "'" + wo.stdJobNo;
                    _cells.GetCell(20, i).Value = "'" + wo.maintSchTask;
                    //PLANNING
                    _cells.GetCell(21, i).Value = "" + wo.autoRequisitionInd;
                    _cells.GetCell(22, i).Value = "" + wo.assignPerson;
                    _cells.GetCell(23, i).Value = "" + wo.planPriority;
                    _cells.GetCell(24, i).Value = "'" + wo.requisitionStartDate;
                    _cells.GetCell(25, i).Value = "'" + wo.requisitionStartTime;
                    _cells.GetCell(26, i).Value = "'" + wo.requiredByDate;
                    _cells.GetCell(27, i).Value = "'" + wo.requiredByTime;
                    _cells.GetCell(28, i).Value = "'" + wo.planStrDate;
                    _cells.GetCell(29, i).Value = "'" + wo.planStrTime;
                    _cells.GetCell(30, i).Value = "'" + wo.planFinDate;
                    _cells.GetCell(31, i).Value = "'" + wo.planFinTime;

                    _cells.GetCell(32, i).Value = "" + wo.unitOfWork;
                    _cells.GetCell(33, i).Value = "" + wo.unitsRequired;
                    _cells.GetCell(34, i).Value = "" + wo.pcComplete + "%/" + wo.unitsComplete;
                    //COST
                    _cells.GetCell(35, i).Value = "" + wo.accountCode;
                    _cells.GetCell(36, i).Value = "" + wo.projectNo;
                    _cells.GetCell(37, i).Value = "'" + wo.parentWo;
                    //JOB_CODES
                    _cells.GetCell(38, i).Value2 = "'" + wo.jobCode1;
                    _cells.GetCell(39, i).Value2 = "'" + wo.jobCode2;
                    _cells.GetCell(40, i).Value2 = "'" + wo.jobCode3;
                    _cells.GetCell(41, i).Value2 = "'" + wo.jobCode4;
                    _cells.GetCell(42, i).Value2 = "'" + wo.jobCode5;
                    _cells.GetCell(43, i).Value = "'" + wo.jobCode6;
                    _cells.GetCell(44, i).Value = "'" + wo.jobCode7;
                    _cells.GetCell(45, i).Value = "'" + wo.jobCode8;
                    _cells.GetCell(46, i).Value = "'" + wo.jobCode9;
                    _cells.GetCell(47, i).Value = "'" + wo.jobCode10;
                    _cells.GetCell(48, i).Value = "'" + wo.locationFr;
                    _cells.GetCell(49, i).Value = "'" + wo.failurePart;
                    _cells.GetCell(50, i).Value = "'" + wo.completedCode;
                    _cells.GetCell(51, i).Value = "" + wo.completeTextFlag;
                    _cells.GetCell(52, i).Value = "" + wo.closeCommitDate;
                    _cells.GetCell(53, i).Value = "" + wo.completedBy;
                    _cells.GetCell(54, i).Value = "" + wo.GetExtendedDescription(urlService, opContext).Header;
                    _cells.GetCell(55, i).Value = "" + wo.GetExtendedDescription(urlService, opContext).Body;
                    _cells.GetRange(54, i, 55, i).WrapText = false;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWODetailedList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void CreateWoDetailedList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameD01, ResultColumnD01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowD01 + 1;
            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDsct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);

            while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)) || !string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value2)))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    //GENERAL
                    wo.districtCode = _frmAuth.EllipseDsct;
                    wo.workGroup = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    string workNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    if (workNo != null)
                    {
                        if (workNo.Length == 2) //prefijo
                            wo.SetWorkOrderDto(workNo, null);
                        else
                            wo.SetWorkOrderDto(workNo);
                    }
                    wo.workOrderStatusM = WoStatusList.GetStatusCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value));
                    wo.workOrderDesc = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                    wo.equipmentNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    wo.compCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                    wo.compModCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    wo.workOrderType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    wo.maintenanceType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value));
                    wo.workOrderStatusU = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                    //DETAILS
                    wo.raisedDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    wo.raisedTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);
                    wo.originatorId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value);
                    wo.origPriority = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value));
                    wo.origDocType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value);
                    wo.origDocNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value);
                    string relatedWo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value);
                    wo.SetRelatedWoDto(relatedWo);
                    wo.requestId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, i).Value);
                    wo.stdJobNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, i).Value);
                    wo.maintSchTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, i).Value);
                    //PLANNING
                    wo.autoRequisitionInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, i).Value);
                    wo.assignPerson = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(22, i).Value);
                    wo.planPriority = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(23, i).Value));
                    wo.requisitionStartDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(24, i).Value);
                    wo.requisitionStartTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(25, i).Value);
                    wo.requiredByDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(26, i).Value);
                    wo.requiredByTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(27, i).Value);
                    wo.planStrDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(28, i).Value);//
                    wo.planStrTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(29, i).Value);//
                    wo.planFinDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(30, i).Value);//
                    wo.planFinTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(31, i).Value);//

                    //Elemento de control para planning
                    if (!planValidation)
                    {
                        wo.planStrDate = null;
                        wo.planStrTime = null;
                        wo.planFinDate = null;
                        wo.planFinTime = null;
                    }
                    //

                    wo.unitOfWork = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(32, i).Value);
                    wo.unitsRequired = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(33, i).Value);
                    //pcComp/unComp
                    //COST
                    wo.accountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(35, i).Value);
                    wo.projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(36, i).Value);
                    wo.parentWo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(37, i).Value);
                    //JOB_CODES
                    wo.jobCode1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(38, i).Value);
                    wo.jobCode2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(39, i).Value);
                    wo.jobCode3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(40, i).Value);
                    wo.jobCode4 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(41, i).Value);
                    wo.jobCode5 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(42, i).Value);
                    wo.jobCode6 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(43, i).Value);
                    wo.jobCode7 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(44, i).Value);
                    wo.jobCode8 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(45, i).Value);
                    wo.jobCode9 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(46, i).Value);
                    wo.jobCode10 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(47, i).Value);
                    wo.locationFr = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(48, i).Value);
                    wo.failurePart = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(49, i).Value);
                    wo.calculatedEquipmentFlag = "true";
                    wo.calculatedMatFlag = "true";
                    wo.calculatedOtherFlag = "true";
                    wo.calculatedLabFlag = "true";
                    wo.calculatedDurationsFlag = "true";//se está forzando porque recientemente en una actualización de E8, si no se envía (se envía nulo) el predeterminado es falso
                    var extendedHeader = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(54, i).Value);
                    var extendedBody = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(55, i).Value);
                    wo.SetExtendedDescription(extendedHeader, extendedBody);

                    List<string> ots = null;
                    if (wo.workGroup == "CALLCEN")
                    {
                        ots = WorkOrderActions.FetchOrigDocNo(_eFunctions, wo.districtCode, wo.workGroup, wo.origDocType, wo.origDocNo);
                        if (ots != null && ots.Count > 0)
                        {
                            _cells.GetCell(ResultColumnD01, i).Value = "YA EXISTE LA ORDEN RELACIONADA " + wo.origDocNo + " DEL SISTEMA MAXIMO";
                            _cells.GetCell(ResultColumnD01, i).Style = StyleConstants.Warning;
                        }

                    }

                    if (ots != null && ots.Count > 0) continue;
                    var replySheet = WorkOrderActions.CreateWorkOrder(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, wo);
                    wo.SetWorkOrderDto(replySheet.workOrder.prefix, replySheet.workOrder.no);
                    ReplyMessage replyExtended = WorkOrderActions.UpdateWorkOrderExtendedDescription(urlService, opSheet, district, wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no, extendedHeader, extendedBody);

                    var stringErrors = "";

                    if (replyExtended != null && replyExtended.Errors != null && replyExtended.Errors.Length > 0)
                        stringErrors = replyExtended.Errors.Aggregate(stringErrors, (current, error) => current + ("/" + error));

                    _cells.GetCell(ResultColumnD01, i).Value = "CREADA " + replySheet.workOrder.prefix + replySheet.workOrder.no + stringErrors;
                    _cells.GetCell(2, i).Value = replySheet.workOrder.prefix + replySheet.workOrder.no;
                    _cells.GetCell(1, i).Style = string.IsNullOrWhiteSpace(stringErrors) ? StyleConstants.Success : StyleConstants.Warning;
                    _cells.GetCell(ResultColumnD01, i).Style = string.IsNullOrWhiteSpace(stringErrors) ? StyleConstants.Success : StyleConstants.Warning;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumnD01, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void UpdateWoDetailedList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNameD01, ResultColumnD01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowD01 + 1;
            const int validationRow = TitleRowD01 - 1;

            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDsct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);


            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    wo.districtCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("B3").Value);

                    //GENERAL
                    wo.workGroup = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    wo.SetWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value));
                    wo.workOrderStatusM = WoStatusList.GetStatusCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value));
                    wo.workOrderDesc = MyUtilities.IsTrue(_cells.GetCell(4, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value) : null;
                    wo.equipmentNo = MyUtilities.IsTrue(_cells.GetCell(5, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value) : null;
                    wo.compCode = MyUtilities.IsTrue(_cells.GetCell(6, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value) : null;
                    wo.compModCode = MyUtilities.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null;
                    wo.workOrderType = MyUtilities.IsTrue(_cells.GetCell(8, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value)) : null;
                    wo.maintenanceType = MyUtilities.IsTrue(_cells.GetCell(9, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)) : null;
                    wo.workOrderStatusU = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value)) : null;
                    //DETAILS
                    wo.raisedDate = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    wo.raisedTime = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;
                    wo.originatorId = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    wo.origPriority = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)) : null;
                    wo.origDocType = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    wo.origDocNo = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    if (MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value))
                        wo.SetRelatedWoDto(_cells.GetEmptyIfNull(_cells.GetCell(17, i).Value));
                    wo.requestId = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;
                    wo.stdJobNo = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null;
                    wo.maintSchTask = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null;
                    //PLANNING
                    wo.autoRequisitionInd = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null;
                    wo.assignPerson = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    wo.planPriority = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(23, i).Value)) : null;
                    wo.requisitionStartDate = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    wo.requisitionStartTime = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null;
                    wo.requiredByDate = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null;
                    wo.requiredByTime = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null;
                    wo.planStrDate = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null;
                    wo.planStrTime = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;
                    wo.planFinDate = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null;
                    wo.planFinTime = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null;
                    wo.unitOfWork = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null;
                    wo.unitsRequired = MyUtilities.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null;
                    //pcComp/ucComp

                    //Elemento de control para planning
                    if (!planValidation)
                    {
                        wo.planStrDate = null;
                        wo.planStrTime = null;
                        wo.planFinDate = null;
                        wo.planFinTime = null;
                    }
                    //

                    //COST
                    wo.accountCode = MyUtilities.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null;
                    wo.projectNo = MyUtilities.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null;
                    wo.parentWo = MyUtilities.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null;
                    //JOB_CODES
                    wo.jobCode1 = MyUtilities.IsTrue(_cells.GetCell(38, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value) : null;
                    wo.jobCode2 = MyUtilities.IsTrue(_cells.GetCell(39, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value) : null;
                    wo.jobCode3 = MyUtilities.IsTrue(_cells.GetCell(40, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(40, i).Value) : null;
                    wo.jobCode4 = MyUtilities.IsTrue(_cells.GetCell(41, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(41, i).Value) : null;
                    wo.jobCode5 = MyUtilities.IsTrue(_cells.GetCell(42, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(42, i).Value) : null;
                    wo.jobCode6 = MyUtilities.IsTrue(_cells.GetCell(43, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(43, i).Value) : null;
                    wo.jobCode7 = MyUtilities.IsTrue(_cells.GetCell(44, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(44, i).Value) : null;
                    wo.jobCode8 = MyUtilities.IsTrue(_cells.GetCell(45, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(45, i).Value) : null;
                    wo.jobCode9 = MyUtilities.IsTrue(_cells.GetCell(46, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(46, i).Value) : null;
                    wo.jobCode10 = MyUtilities.IsTrue(_cells.GetCell(47, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(47, i).Value) : null;
                    wo.locationFr = MyUtilities.IsTrue(_cells.GetCell(48, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(48, i).Value) : null;
                    wo.failurePart = MyUtilities.IsTrue(_cells.GetCell(49, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(49, i).Value) : null;

                    //wo.calculatedEquipmentFlag = "true";
                    //wo.calculatedMatFlag = "true";
                    //wo.calculatedOtherFlag = "true";
                    //wo.calculatedLabFlag = "true";
                    //wo.calculatedDurationsFlag = "true";
                    var extendedHeader = MyUtilities.IsTrue(_cells.GetCell(54, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(54, i).Value) : null;
                    var extendedBody = MyUtilities.IsTrue(_cells.GetCell(55, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(55, i).Value) : null;
                    wo.SetExtendedDescription(extendedHeader, extendedBody);

                    WorkOrderActions.ModifyWorkOrder(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, wo);
                    var replyExtended = WorkOrderActions.UpdateWorkOrderExtendedDescription(urlService, opSheet, district, wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no, extendedHeader, extendedBody);
                    var stringErrors = "";

                    if (replyExtended != null && replyExtended.Errors != null && replyExtended.Errors.Length > 0)
                        foreach (var error in replyExtended.Errors)
                            stringErrors += "/" + error;


                    _cells.GetCell(ResultColumnD01, i).Value = "ACTUALIZADA " + stringErrors;
                    _cells.GetCell(1, i).Style = string.IsNullOrWhiteSpace(stringErrors) ? StyleConstants.Success : StyleConstants.Warning;
                    _cells.GetCell(ResultColumnD01, i).Style = string.IsNullOrWhiteSpace(stringErrors) ? StyleConstants.Success : StyleConstants.Warning;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumnD01, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void FormatQuality()
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
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 4
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameQ01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "WORK ORDERS - ELLIPSE 8";
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
                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var statusList = WoStatusList.GetStatusNames(true);
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.GetCell("A4").AddComment("--ÁREA/SUPERINTENDENCIA--\n" +
                    "INST: IMIS, MINA\n" +
                    "MDC: FFCC, PBV, PTAS\n" +
                    "MNTTO: MINA\n" +
                    "SOP: ENERGIA, LIVIANOS, MEDIANOS, GRUAS, ENERGIA");
                _cells.GetCell("A4").Comment.Shape.TextFrame.AutoSize = true;

                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), workGroupList, ValidationSheetName, 3, false);
                _cells.GetCell("A5").Value = SearchFieldCriteriaType.EquipmentReference.Value;
                _cells.SetValidationList(_cells.GetCell("A5"), ValidationSheetName, 2);
                _cells.GetCell("A6").Value = "STATUS";
                _cells.SetValidationList(_cells.GetCell("B6"), statusList, ValidationSheetName, 4);
                _cells.GetRange("A3", "A6").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B6").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.Raised.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 5);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);


                _cells.GetRange(1, TitleRowQ01, ResultColumnQ01, TitleRowQ01).Style = StyleConstants.TitleRequired;

                //GENERAL

                _cells.GetCell(01, TitleRowQ01).Value = "WORK_GROUP";
                _cells.GetCell(02, TitleRowQ01).Value = "WORK_ORDER";
                _cells.GetCell(03, TitleRowQ01).Value = "WO_STATUS";
                _cells.GetCell(04, TitleRowQ01).Value = "DESCRIPTION";
                _cells.GetCell(05, TitleRowQ01).Value = "EQUIPMENT";
                _cells.GetCell(06, TitleRowQ01).Value = "COMP_CODE";
                _cells.GetCell(07, TitleRowQ01).Value = "MOD_CODE";
                _cells.GetCell(08, TitleRowQ01).Value = "WO_TYPE";
                _cells.GetCell(09, TitleRowQ01).Value = "MT_TYPE";
                _cells.GetCell(10, TitleRowQ01).Value = "WO_USER_STATUS";
                _cells.GetCell(11, TitleRowQ01).Value = "RAISED_DATE";
                _cells.GetCell(12, TitleRowQ01).Value = "ORIGINATOR_ID";
                _cells.GetCell(13, TitleRowQ01).Value = "ORIG_PRIORITY";
                _cells.GetCell(14, TitleRowQ01).Value = "PLAN_PRIORITY";
                _cells.GetCell(15, TitleRowQ01).Value = "STD_JOB";
                _cells.GetCell(16, TitleRowQ01).Value = "MST";
                _cells.GetCell(17, TitleRowQ01).Value = "PLAN_STR_DATE";
                _cells.GetCell(18, TitleRowQ01).Value = "UNIT_OF_WORK";
                _cells.GetCell(19, TitleRowQ01).Value = "UNITS_REQUIRED";
                _cells.GetCell(20, TitleRowQ01).Value = "PC/UN COMPLETED";
                _cells.GetCell(21, TitleRowQ01).Value = "DUR EST";
                _cells.GetCell(22, TitleRowQ01).Value = "DUR ACT";
                _cells.GetCell(23, TitleRowQ01).Value = "LAB H. EST";
                _cells.GetCell(24, TitleRowQ01).Value = "LAB H. ACT";
                _cells.GetCell(25, TitleRowQ01).Value = "LAB C. EST";
                _cells.GetCell(26, TitleRowQ01).Value = "LAB C. ACT";
                _cells.GetCell(27, TitleRowQ01).Value = "MAT C. EST";
                _cells.GetCell(28, TitleRowQ01).Value = "MAT C. ACT";
                _cells.GetCell(29, TitleRowQ01).Value = "OTH C. EST";
                _cells.GetCell(30, TitleRowQ01).Value = "OTH C. ACT";
                _cells.GetCell(31, TitleRowQ01).Value = "JOBCODES";
                _cells.GetCell(32, TitleRowQ01).Value = "COMPL_DATE";
                _cells.GetCell(33, TitleRowQ01).Value = "COMPL_COD";
                _cells.GetCell(34, TitleRowQ01).Value = "COMP_COMM";
                _cells.GetCell(35, TitleRowQ01).Value = "COMP_BY";
                _cells.GetCell(ResultColumnQ01, TitleRowQ01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnQ01, TitleRowQ01).Style = StyleConstants.TitleResult;

                _cells.FormatAsTable(_cells.GetRange(1, TitleRowQ01, ResultColumnQ01, TitleRowQ01 + 1), TableNameQ01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatQuality()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        private void ReviewWoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            //Obtengo los valores de las opciones de búsqueda
            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A5").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B6").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            var listwo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
            var i = TitleRow01 + 1;
            foreach (var wo in listwo)
            {
                try
                {
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "" + wo.workGroup;
                    _cells.GetCell(2, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(3, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(4, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(5, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(6, i).Value = "'" + wo.compCode;
                    _cells.GetCell(7, i).Value = "'" + wo.compModCode;
                    _cells.GetCell(8, i).Value = "" + wo.workOrderType;
                    _cells.GetCell(9, i).Value = "" + wo.maintenanceType;
                    _cells.GetCell(10, i).Value = "" + wo.workOrderStatusU;
                    //DETAILS
                    _cells.GetCell(11, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(12, i).Value = "'" + wo.raisedTime;
                    _cells.GetCell(13, i).Value = "" + wo.originatorId;
                    _cells.GetCell(14, i).Value = "" + wo.origPriority;
                    _cells.GetCell(15, i).Value = "" + wo.origDocType;
                    _cells.GetCell(16, i).Value = "'" + wo.origDocNo;
                    _cells.GetCell(17, i).Value = "'" + wo.GetRelatedWoDto().prefix + wo.GetRelatedWoDto().no;
                    _cells.GetCell(18, i).Value = "'" + wo.requestId;
                    _cells.GetCell(19, i).Value = "'" + wo.stdJobNo;
                    _cells.GetCell(20, i).Value = "'" + wo.maintSchTask;
                    //PLANNING
                    _cells.GetCell(21, i).Value = "" + wo.autoRequisitionInd;
                    _cells.GetCell(22, i).Value = "" + wo.assignPerson;
                    _cells.GetCell(23, i).Value = "" + wo.planPriority;
                    _cells.GetCell(24, i).Value = "'" + wo.requisitionStartDate;
                    _cells.GetCell(25, i).Value = "'" + wo.requisitionStartTime;
                    _cells.GetCell(26, i).Value = "'" + wo.requiredByDate;
                    _cells.GetCell(27, i).Value = "'" + wo.requiredByTime;
                    _cells.GetCell(28, i).Value = "'" + wo.planStrDate;
                    _cells.GetCell(29, i).Value = "'" + wo.planStrTime;
                    _cells.GetCell(30, i).Value = "'" + wo.planFinDate;
                    _cells.GetCell(31, i).Value = "'" + wo.planFinTime;

                    _cells.GetCell(32, i).Value = "" + wo.unitOfWork;
                    _cells.GetCell(33, i).Value = "" + wo.unitsRequired;
                    _cells.GetCell(34, i).Value = "" + wo.pcComplete + "%/" + wo.unitsComplete;
                    //COST
                    _cells.GetCell(35, i).Value = "" + wo.accountCode;
                    _cells.GetCell(36, i).Value = "" + wo.projectNo;
                    _cells.GetCell(37, i).Value = "'" + wo.parentWo;
                    //JOB_CODES
                    _cells.GetCell(38, i).Value2 = "'" + wo.jobCode1;
                    _cells.GetCell(39, i).Value2 = "'" + wo.jobCode2;
                    _cells.GetCell(40, i).Value2 = "'" + wo.jobCode3;
                    _cells.GetCell(41, i).Value2 = "'" + wo.jobCode4;
                    _cells.GetCell(42, i).Value2 = "'" + wo.jobCode5;
                    _cells.GetCell(43, i).Value = "'" + wo.jobCode6;
                    _cells.GetCell(44, i).Value = "'" + wo.jobCode7;
                    _cells.GetCell(45, i).Value = "'" + wo.jobCode8;
                    _cells.GetCell(46, i).Value = "'" + wo.jobCode9;
                    _cells.GetCell(47, i).Value = "'" + wo.jobCode10;
                    _cells.GetCell(48, i).Value = "'" + wo.locationFr;
                    _cells.GetCell(49, i).Value = "'" + wo.failurePart;
                    _cells.GetCell(50, i).Value = "'" + wo.completedCode;
                    _cells.GetCell(51, i).Value = "" + wo.completeTextFlag;
                    _cells.GetCell(52, i).Value = "" + wo.closeCommitDate;
                    _cells.GetCell(53, i).Value = "" + wo.completedBy;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", ex.Message);
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
        private void ReReviewWoList()
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
                    var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                    district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;
                    var woNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, woNo);

                    if (wo == null || wo.GetWorkOrderDto().no == null)
                        throw new Exception("WORK ORDER NO ENCONTRADA");
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "" + wo.workGroup;
                    _cells.GetCell(2, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(3, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(4, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(5, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(6, i).Value = "'" + wo.compCode;
                    _cells.GetCell(7, i).Value = "'" + wo.compModCode;
                    _cells.GetCell(8, i).Value = "" + wo.workOrderType;
                    _cells.GetCell(9, i).Value = "" + wo.maintenanceType;
                    _cells.GetCell(10, i).Value = "" + wo.workOrderStatusU;
                    //DETAILS
                    _cells.GetCell(11, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(12, i).Value = "'" + wo.raisedTime;
                    _cells.GetCell(13, i).Value = "" + wo.originatorId;
                    _cells.GetCell(14, i).Value = "" + wo.origPriority;
                    _cells.GetCell(15, i).Value = "" + wo.origDocType;
                    _cells.GetCell(16, i).Value = "'" + wo.origDocNo;
                    _cells.GetCell(17, i).Value = "'" + wo.GetRelatedWoDto().prefix + wo.GetRelatedWoDto().no;
                    _cells.GetCell(18, i).Value = "'" + wo.requestId;
                    _cells.GetCell(19, i).Value = "'" + wo.stdJobNo;
                    _cells.GetCell(20, i).Value = "'" + wo.maintSchTask;
                    //PLANNING
                    _cells.GetCell(21, i).Value = "" + wo.autoRequisitionInd;
                    _cells.GetCell(22, i).Value = "" + wo.assignPerson;
                    _cells.GetCell(23, i).Value = "" + wo.planPriority;
                    _cells.GetCell(24, i).Value = "'" + wo.requisitionStartDate;
                    _cells.GetCell(25, i).Value = "'" + wo.requisitionStartTime;
                    _cells.GetCell(26, i).Value = "'" + wo.requiredByDate;
                    _cells.GetCell(27, i).Value = "'" + wo.requiredByTime;
                    _cells.GetCell(28, i).Value = "'" + wo.planStrDate;
                    _cells.GetCell(29, i).Value = "'" + wo.planStrTime;
                    _cells.GetCell(30, i).Value = "'" + wo.planFinDate;
                    _cells.GetCell(31, i).Value = "'" + wo.planFinTime;

                    _cells.GetCell(32, i).Value = "" + wo.unitOfWork;
                    _cells.GetCell(33, i).Value = "" + wo.unitsRequired;
                    _cells.GetCell(34, i).Value = "" + wo.pcComplete + "%/" + wo.unitsComplete;
                    //COST
                    _cells.GetCell(35, i).Value = "" + wo.accountCode;
                    _cells.GetCell(36, i).Value = "" + wo.projectNo;
                    _cells.GetCell(37, i).Value = "'" + wo.parentWo;
                    //JOB_CODES
                    _cells.GetCell(38, i).Value2 = "'" + wo.jobCode1;
                    _cells.GetCell(39, i).Value2 = "'" + wo.jobCode2;
                    _cells.GetCell(40, i).Value2 = "'" + wo.jobCode3;
                    _cells.GetCell(41, i).Value2 = "'" + wo.jobCode4;
                    _cells.GetCell(42, i).Value2 = "'" + wo.jobCode5;
                    _cells.GetCell(43, i).Value = "'" + wo.jobCode6;
                    _cells.GetCell(44, i).Value = "'" + wo.jobCode7;
                    _cells.GetCell(45, i).Value = "'" + wo.jobCode8;
                    _cells.GetCell(46, i).Value = "'" + wo.jobCode9;
                    _cells.GetCell(47, i).Value = "'" + wo.jobCode10;
                    _cells.GetCell(48, i).Value = "'" + wo.locationFr;
                    _cells.GetCell(49, i).Value = "'" + wo.failurePart;
                    _cells.GetCell(50, i).Value = "'" + wo.completedCode;
                    _cells.GetCell(51, i).Value = "" + wo.completeTextFlag;
                    _cells.GetCell(52, i).Value = "" + wo.closeCommitDate;
                    _cells.GetCell(53, i).Value = "" + wo.completedBy;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReviewQualityList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableNameQ01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes();

            //Obtengo los valores de las opciones de búsqueda
            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A5").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);
            var statusKey = _cells.GetEmptyIfNull(_cells.GetCell("B6").Value);
            var dateCriteriaKeyText = _cells.GetEmptyIfNull(_cells.GetCell("D3").Value);
            var startDate = _cells.GetEmptyIfNull(_cells.GetCell("D4").Value);
            var endDate = _cells.GetEmptyIfNull(_cells.GetCell("D5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var dateCriteriaKey = dateCriteriaList.FirstOrDefault(v => v.Value.Equals(dateCriteriaKeyText)).Key;

            var completeCodeList = _eFunctions.GetDictionaryItemCodes("SC");

            var listwo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
            var i = TitleRowQ01 + 1;
            foreach (var wo in listwo)
            {
                try
                {
                    var wc = new WorkOrderQualityStyles(wo);
                    //GENERAL
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumnQ01, i).Style = StyleConstants.Normal;
                    _cells.GetCell(1, i).Value = "" + wo.workGroup;
                    _cells.GetCell(1, i).Style = wc.WorkGroup;
                    _cells.GetCell(2, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(3, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(3, i).Style = wc.WorkOrderStatusM;
                    _cells.GetCell(4, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(5, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(5, i).Style = wc.EquipmentNo;
                    _cells.GetCell(6, i).Value = "" + wo.compCode;
                    _cells.GetCell(6, i).Style = wc.CompCode;
                    _cells.GetCell(7, i).Value = "" + wo.compModCode;
                    _cells.GetCell(8, i).Value = "" + wo.workOrderType;
                    _cells.GetCell(8, i).Style = wc.WorkOrderType;
                    _cells.GetCell(9, i).Value = "" + wo.maintenanceType;
                    _cells.GetCell(9, i).Style = wc.MaintenanceType;
                    _cells.GetCell(10, i).Value = "" + wo.workOrderStatusU;
                    _cells.GetCell(10, i).Style = wc.WorkOrderStatusU;
                    //DETAILS
                    _cells.GetCell(11, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(12, i).Value = "" + wo.originatorId;
                    _cells.GetCell(13, i).Value = "" + wo.origPriority;
                    _cells.GetCell(13, i).Style = wc.OriginatorPriority;
                    _cells.GetCell(14, i).Value = "" + wo.planPriority;
                    _cells.GetCell(14, i).Style = wc.PlanPriority;
                    _cells.GetCell(15, i).Value = "" + wo.stdJobNo;
                    _cells.GetCell(16, i).Value = "" + wo.maintSchTask;
                    //PLANNING
                    _cells.GetCell(17, i).Value = "'" + wo.planStrDate;
                    _cells.GetCell(18, i).Value = "" + wo.unitOfWork;
                    _cells.GetCell(18, i).Style = wc.UnitOfWork;
                    _cells.GetCell(19, i).Value = "" + wo.unitsRequired;
                    _cells.GetCell(19, i).Style = wc.UnitsRequired;
                    _cells.GetCell(20, i).Value = "" + wo.pcComplete + "%/" + wo.unitsComplete;
                    _cells.GetCell(20, i).Style = wc.UnitsCompleted;
                    //ESTIMATES
                    _cells.GetCell(21, i).Value = "" + wo.estimatedDurationsHrs;
                    _cells.GetCell(22, i).Value = "" + wo.actualDurationsHrs;
                    _cells.GetCell(22, i).Style = wc.ActualDurationHrs;
                    _cells.GetCell(23, i).Value = "" + (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabHrs : wo.estimatedLabHrs);
                    _cells.GetCell(24, i).Value2 = "" + wo.actualLabHrs;
                    _cells.GetCell(24, i).Style = wc.ActualLabHrs;
                    _cells.GetCell(25, i).Value2 = "" + (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabCost : wo.estimatedLabCost);
                    _cells.GetCell(26, i).Value2 = "" + wo.actualLabCost;
                    _cells.GetCell(26, i).Style = wc.ActualLabCost;
                    _cells.GetCell(27, i).Value2 = "" + (wo.calculatedMatFlag.Equals("Y") ? wo.calculatedMatCost : wo.estimatedMatCost);
                    _cells.GetCell(28, i).Value2 = "" + wo.actualMatCost;
                    _cells.GetCell(28, i).Style = wc.ActualMatCost;
                    _cells.GetCell(29, i).Value = "" + wo.estimatedOtherCost;
                    _cells.GetCell(30, i).Value = "" + wo.actualOtherCost;
                    _cells.GetCell(30, i).Style = wc.ActualOtherCost;
                    //
                    _cells.GetCell(31, i).Value = "" + wo.jobCodeFlag;
                    _cells.GetCell(31, i).Style = wc.JobCodesFlag;
                    _cells.GetCell(32, i).Value = "'" + wo.closeCommitDate;
                    string outString;
                    _cells.GetCell(33, i).Value = completeCodeList.ContainsKey(wo.completedCode) ? "'" + wo.completedCode + " - " + completeCodeList.TryGetValue(wo.completedCode, out outString) : "";
                    _cells.GetCell(34, i).Value = "" + wo.completeTextFlag;
                    _cells.GetCell(34, i).Style = wc.CompleteTextFlag;
                    _cells.GetCell(35, i).Value = "" + wo.completedBy;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnQ01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewQualityList()", ex.Message);
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
        private void ReReviewQualityList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var completeCodeList = _eFunctions.GetDictionaryItemCodes("SC");
            var i = TitleRowQ01 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var woNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, "" + _cells.GetCell("B3").Value, woNo);
                    if (wo == null || wo.GetWorkOrderDto().no == null)
                        throw new Exception("WORK ORDER NO ENCONTRADA");

                    var wc = new WorkOrderQualityStyles(wo);
                    //GENERAL
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumnQ01, i).Style = StyleConstants.Normal;
                    _cells.GetCell(1, i).Value = "" + wo.workGroup;
                    _cells.GetCell(1, i).Style = wc.WorkGroup;
                    _cells.GetCell(2, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(3, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(3, i).Style = wc.WorkOrderStatusM;
                    _cells.GetCell(4, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(5, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(5, i).Style = wc.EquipmentNo;
                    _cells.GetCell(6, i).Value = "" + wo.compCode;
                    _cells.GetCell(6, i).Style = wc.CompCode;
                    _cells.GetCell(7, i).Value = "" + wo.compModCode;
                    _cells.GetCell(8, i).Value = "" + wo.workOrderType;
                    _cells.GetCell(8, i).Style = wc.WorkOrderType;
                    _cells.GetCell(9, i).Value = "" + wo.maintenanceType;
                    _cells.GetCell(9, i).Style = wc.MaintenanceType;
                    _cells.GetCell(10, i).Value = "" + wo.workOrderStatusU;
                    _cells.GetCell(10, i).Style = wc.WorkOrderStatusU;
                    //DETAILS
                    _cells.GetCell(11, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(12, i).Value = "" + wo.originatorId;
                    _cells.GetCell(13, i).Value = "" + wo.origPriority;
                    _cells.GetCell(13, i).Style = wc.OriginatorPriority;
                    _cells.GetCell(14, i).Value = "" + wo.planPriority;
                    _cells.GetCell(14, i).Style = wc.PlanPriority;
                    _cells.GetCell(15, i).Value = "" + wo.stdJobNo;
                    _cells.GetCell(16, i).Value = "" + wo.maintSchTask;
                    //PLANNING
                    _cells.GetCell(17, i).Value = "'" + wo.planStrDate;
                    _cells.GetCell(18, i).Value = "" + wo.unitOfWork;
                    _cells.GetCell(18, i).Style = wc.UnitOfWork;
                    _cells.GetCell(19, i).Value = "" + wo.unitsRequired;
                    _cells.GetCell(19, i).Style = wc.UnitsRequired;
                    _cells.GetCell(20, i).Value = "" + wo.pcComplete + "%/" + wo.unitsComplete;
                    _cells.GetCell(20, i).Style = wc.UnitsCompleted;
                    //ESTIMATES
                    _cells.GetCell(21, i).Value = "" + wo.estimatedDurationsHrs;
                    _cells.GetCell(22, i).Value = "" + wo.actualDurationsHrs;
                    _cells.GetCell(22, i).Style = wc.ActualDurationHrs;
                    _cells.GetCell(23, i).Value = "" + (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabHrs : wo.estimatedLabHrs);
                    _cells.GetCell(24, i).Value2 = "" + wo.actualLabHrs;
                    _cells.GetCell(24, i).Style = wc.ActualLabHrs;
                    _cells.GetCell(25, i).Value2 = "" + (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabCost : wo.estimatedLabCost);
                    _cells.GetCell(26, i).Value2 = "" + wo.actualLabCost;
                    _cells.GetCell(26, i).Style = wc.ActualLabCost;
                    _cells.GetCell(27, i).Value2 = "" + (wo.calculatedMatFlag.Equals("Y") ? wo.calculatedMatCost : wo.estimatedMatCost);
                    _cells.GetCell(28, i).Value2 = "" + wo.actualMatCost;
                    _cells.GetCell(28, i).Style = wc.ActualMatCost;
                    _cells.GetCell(29, i).Value = "" + wo.estimatedOtherCost;
                    _cells.GetCell(30, i).Value = "" + wo.actualOtherCost;
                    _cells.GetCell(30, i).Style = wc.ActualOtherCost;
                    //
                    _cells.GetCell(31, i).Value = "" + wo.jobCodeFlag;
                    _cells.GetCell(31, i).Style = wc.JobCodesFlag;
                    _cells.GetCell(32, i).Value = "'" + wo.closeCommitDate;
                    string outString;

                    _cells.GetCell(33, i).Value = completeCodeList.ContainsKey(wo.completedCode) ? "'" + wo.completedCode + " - " + completeCodeList.TryGetValue(wo.completedCode, out outString) : "";
                    _cells.GetCell(34, i).Value = "" + wo.completeTextFlag;
                    _cells.GetCell(34, i).Style = wc.CompleteTextFlag;
                    _cells.GetCell(35, i).Value = "" + wo.completedBy;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnQ01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void CreateWoList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDsct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);

            while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)) || !string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value2)))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    //GENERAL
                    wo.workGroup = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    string workNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    if (workNo != null)
                    {
                        if (workNo.Length == 2) //prefijo
                            wo.SetWorkOrderDto(workNo, null);
                        else
                            wo.SetWorkOrderDto(workNo);
                    }
                    wo.workOrderStatusM = WoStatusList.GetStatusCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value));
                    wo.workOrderDesc = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                    wo.equipmentNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    wo.compCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                    wo.compModCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    wo.workOrderType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    wo.maintenanceType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value));
                    wo.workOrderStatusU = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                    //DETAILS
                    wo.raisedDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    wo.raisedTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);
                    wo.originatorId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value);
                    wo.origPriority = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value));
                    wo.origDocType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value);
                    wo.origDocNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value);
                    string relatedWo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value);
                    wo.SetRelatedWoDto(relatedWo);
                    wo.requestId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, i).Value);
                    wo.stdJobNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, i).Value);
                    wo.maintSchTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, i).Value);
                    //PLANNING
                    wo.autoRequisitionInd = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, i).Value);
                    wo.assignPerson = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(22, i).Value);
                    wo.planPriority = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(23, i).Value));
                    wo.requisitionStartDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(24, i).Value);
                    wo.requisitionStartTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(25, i).Value);
                    wo.requiredByDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(26, i).Value);
                    wo.requiredByTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(27, i).Value);
                    wo.planStrDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(28, i).Value);//
                    wo.planStrTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(29, i).Value);//
                    wo.planFinDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(30, i).Value);//
                    wo.planFinTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(31, i).Value);//

                    //Elemento de control para planning
                    if (!planValidation)
                    {
                        wo.planStrDate = null;
                        wo.planStrTime = null;
                        wo.planFinDate = null;
                        wo.planFinTime = null;
                    }
                    //

                    wo.unitOfWork = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(32, i).Value);
                    wo.unitsRequired = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(33, i).Value);
                    //pcComp/unComp
                    //COST
                    wo.accountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(35, i).Value);
                    wo.projectNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(36, i).Value);
                    wo.parentWo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(37, i).Value);
                    //JOB_CODES
                    wo.jobCode1 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(38, i).Value);
                    wo.jobCode2 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(39, i).Value);
                    wo.jobCode3 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(40, i).Value);
                    wo.jobCode4 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(41, i).Value);
                    wo.jobCode5 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(42, i).Value);
                    wo.jobCode6 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(43, i).Value);
                    wo.jobCode7 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(44, i).Value);
                    wo.jobCode8 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(45, i).Value);
                    wo.jobCode9 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(46, i).Value);
                    wo.jobCode10 = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(47, i).Value);
                    wo.locationFr = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(48, i).Value);
                    wo.failurePart = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(49, i).Value);
                    wo.calculatedEquipmentFlag = "true";
                    wo.calculatedMatFlag = "true";
                    wo.calculatedOtherFlag = "true";
                    wo.calculatedLabFlag = "true";
                    wo.calculatedDurationsFlag = "true";//se está forzando porque recientemente en una actualización de E8, si no se envía (se envía nulo) el predeterminado es falso

                    var replySheet = WorkOrderActions.CreateWorkOrder(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, wo);

                    _cells.GetCell(ResultColumn01, i).Value = "CREADA " + replySheet.workOrder.prefix + replySheet.workOrder.no;
                    _cells.GetCell(2, i).Value = replySheet.workOrder.prefix + replySheet.workOrder.no;
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void UpdateWoList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDsct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnviroment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);


            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    wo.districtCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("B3").Value);

                    //GENERAL
                    wo.workGroup = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    wo.SetWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value));
                    wo.workOrderStatusM = WoStatusList.GetStatusCode(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value));
                    wo.workOrderDesc = MyUtilities.IsTrue(_cells.GetCell(4, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value) : null;
                    wo.equipmentNo = MyUtilities.IsTrue(_cells.GetCell(5, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value) : null;
                    wo.compCode = MyUtilities.IsTrue(_cells.GetCell(6, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value) : null;
                    wo.compModCode = MyUtilities.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null;
                    wo.workOrderType = MyUtilities.IsTrue(_cells.GetCell(8, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(8, i).Value)) : null;
                    wo.maintenanceType = MyUtilities.IsTrue(_cells.GetCell(9, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)) : null;
                    wo.workOrderStatusU = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value)) : null;
                    //DETAILS
                    wo.raisedDate = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    wo.raisedTime = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;
                    wo.originatorId = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    wo.origPriority = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)) : null;
                    wo.origDocType = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    wo.origDocNo = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    if (MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value))
                        wo.SetRelatedWoDto(_cells.GetEmptyIfNull(_cells.GetCell(17, i).Value));
                    wo.requestId = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;
                    wo.stdJobNo = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null;
                    wo.maintSchTask = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null;
                    //PLANNING
                    wo.autoRequisitionInd = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null;
                    wo.assignPerson = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    wo.planPriority = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(23, i).Value)) : null;
                    wo.requisitionStartDate = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    wo.requisitionStartTime = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null;
                    wo.requiredByDate = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null;
                    wo.requiredByTime = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null;
                    wo.planStrDate = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null;
                    wo.planStrTime = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;
                    wo.planFinDate = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null;
                    wo.planFinTime = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null;
                    wo.unitOfWork = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null;
                    wo.unitsRequired = MyUtilities.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null;
                    
                    //pcComp/ucComp

                    //Elemento de control para planning
                    if (!planValidation)
                    {
                        wo.planStrDate = null;
                        wo.planStrTime = null;
                        wo.planFinDate = null;
                        wo.planFinTime = null;
                    }
                    //

                    //COST
                    wo.accountCode = MyUtilities.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null;
                    wo.projectNo = MyUtilities.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null;
                    wo.parentWo = MyUtilities.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null;
                    //JOB_CODES
                    wo.jobCode1 = MyUtilities.IsTrue(_cells.GetCell(38, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value) : null;
                    wo.jobCode2 = MyUtilities.IsTrue(_cells.GetCell(39, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value) : null;
                    wo.jobCode3 = MyUtilities.IsTrue(_cells.GetCell(40, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(40, i).Value) : null;
                    wo.jobCode4 = MyUtilities.IsTrue(_cells.GetCell(41, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(41, i).Value) : null;
                    wo.jobCode5 = MyUtilities.IsTrue(_cells.GetCell(42, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(42, i).Value) : null;
                    wo.jobCode6 = MyUtilities.IsTrue(_cells.GetCell(43, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(43, i).Value) : null;
                    wo.jobCode7 = MyUtilities.IsTrue(_cells.GetCell(44, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(44, i).Value) : null;
                    wo.jobCode8 = MyUtilities.IsTrue(_cells.GetCell(45, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(45, i).Value) : null;
                    wo.jobCode9 = MyUtilities.IsTrue(_cells.GetCell(46, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(46, i).Value) : null;
                    wo.jobCode10 = MyUtilities.IsTrue(_cells.GetCell(47, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(47, i).Value) : null;
                    wo.locationFr = MyUtilities.IsTrue(_cells.GetCell(48, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(48, i).Value) : null;
                    wo.failurePart = MyUtilities.IsTrue(_cells.GetCell(49, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(49, i).Value) : null;

                    //wo.calculatedEquipmentFlag = "true";
                    //wo.calculatedMatFlag = "true";
                    //wo.calculatedOtherFlag = "true";
                    //wo.calculatedLabFlag = "true";
                    //wo.calculatedDurationsFlag = "true";

                    WorkOrderActions.ModifyWorkOrder(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, wo);

                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADA";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void CompleteWoList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName04, ResultColumn04);

            var i = TitleRow04 + 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var ignoreClosedStatus = cbIgnoreClosedStatus.Checked;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrderCompleteAtributes();
                    //GENERAL
                    wo.districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
                    wo.workOrder = WorkOrderActions.GetNewWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    wo.closedDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    wo.closedTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    wo.completedBy = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                    wo.completedCode = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value));
                    wo.outServDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                    wo.completeCommentToAppend = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    if (!ignoreClosedStatus)
                    {
                        var woData = WorkOrderActions.FetchWorkOrder(_eFunctions, wo.districtCode, wo.workOrder.prefix + wo.workOrder.no);
                        _eFunctions.CloseConnection();
                        if (WoStatusList.ClosedCode.Equals(woData.workOrderStatusM.Trim()) || WoStatusList.CancelledCode.Equals(woData.workOrderStatusM.Trim()))
                            throw new Exception("La orden " + wo.workOrder.prefix + wo.workOrder.no + " ya está cerrada como " + WoStatusList.GetStatusName(woData.workOrderStatusM.Trim()) + " con código " + woData.completedCode);
                    }
                    var reply = WorkOrderActions.CompleteWorkOrder(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, wo);
                    if (reply.completedCode.Trim() == wo.completedCode.Trim() && reply.closedDate == wo.closedDate)
                    {
                        _cells.GetCell(ResultColumn04, i).Value = "COMPLETADA";
                        _cells.GetCell(1, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumn04, i).Style = StyleConstants.Success;

                        if (!string.IsNullOrWhiteSpace(wo.completeCommentToAppend)) continue;
                        _cells.GetCell(ResultColumn04, i).Value = "COMPLETADA / No se han ingresado comentarios";
                        _cells.GetCell(ResultColumn04, i).Style = StyleConstants.Warning;
                    }
                    else
                    {
                        _cells.GetCell(ResultColumn04, i).Value = "NO SE REALIZÓ ACCIÓN";
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn04, i).Style = StyleConstants.Error;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn04, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn04, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CompleteWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn04, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReOpenWoList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName04, ResultColumn04);

            var i = TitleRow04 + 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var ignoreClosedStatus = cbIgnoreClosedStatus.Checked;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    //GENERAL
                    wo.districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
                    wo.SetWorkOrderDto(WorkOrderActions.GetNewWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value)));
                    if (!ignoreClosedStatus)
                    {
                        var woData = WorkOrderActions.FetchWorkOrder(_eFunctions, wo.districtCode, wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no);
                        _eFunctions.CloseConnection();
                        if (!WoStatusList.ClosedCode.Equals(woData.workOrderStatusM.Trim()) && !WoStatusList.CancelledCode.Equals(woData.workOrderStatusM.Trim()))
                            throw new Exception("La orden " + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no + " ya está abierta como " + WoStatusList.GetStatusName(woData.workOrderStatusM.Trim()));
                    }
                    WorkOrderActions.ReOpenWorkOrder(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, wo);

                    _cells.GetCell(ResultColumn04, i).Value = "REABIERTA";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn04, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn04, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn04, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReOpenWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn04, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReviewCloseText()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName05, ResultColumn05);

            var i = TitleRow05 + 1;

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    //GENERAL
                    var wo = WorkOrderActions.GetNewWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    string districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
                    var closeText = WorkOrderActions.GetWorkOrderCloseText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), districtCode, _frmAuth.EllipsePost, Debugger.DebugWarnings, wo);

                    _cells.GetCell(ResultColumn05 - 1, i).Value = closeText;
                    _cells.GetCell(ResultColumn05, i).Value = "CONSULTA";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05 - 1, i).Value = "";
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewCloseText()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void UpdateCloseText()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName05, ResultColumn05);

            var i = TitleRow05 + 1;

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDsct, _frmAuth.EllipsePost);
            var opSheet = new WorkOrderService.OperationContext
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
                    //GENERAL
                    var wo = WorkOrderActions.GetNewWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    var closeText = _cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value2);
                    var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value);
                    //WorkOrderActions.SetWorkOrderCloseText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), districtCode, _frmAuth.EllipsePost, Debugger.DebugWarnings, wo, closeText);
                    WorkOrderActions.AppendTextToCloseComment(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, districtCode, wo.prefix + wo.no, closeText);

                    _cells.GetCell(ResultColumn05, i).Value = "ACTUALIZADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateCloseText()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn05, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void GetDurationWoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var woCell = new ExcelStyleCells(_excelApp, SheetName01);
            var districtCode = woCell.GetEmptyIfNull(woCell.GetCell("B3").Value);
            _cells.ClearTableRange(TableName06);

            if (_cells.GetNullIfTrimmedEmpty(districtCode) != null)
            {
                _cells.ClearTableRange(TableName06);
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var opSheet = new WorkOrderService.OperationContext
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
                var k = TitleRow06 + 1;
                var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

                while (!string.IsNullOrWhiteSpace(_cells.GetNullOrTrimmedValue(woCell.GetCell(2, i).Value)))
                {
                    try
                    {
                        var wo = WorkOrderActions.GetNewWorkOrderDto(woCell.GetEmptyIfNull(woCell.GetCell(2, i).Value));
                        var durations = WorkOrderActions.GetWorkOrderDurations(urlService, opSheet, districtCode, wo);

                        foreach (var dur in durations)
                        {
                            _cells.GetCell(1, k).Value = districtCode;
                            _cells.GetCell(2, k).Value = wo.prefix + wo.no;
                            _cells.GetCell(3, k).Value = "'" + dur.jobDurationsDate;
                            _cells.GetCell(4, k).Value = "'" + dur.jobDurationsCode;
                            _cells.GetCell(5, k).Value = "'" + dur.jobDurationsStart;
                            _cells.GetCell(6, k).Value = "'" + dur.jobDurationsFinish;
                            _cells.GetCell(7, k).Value = "";//TO DO Add validation
                            k++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, k).Value = districtCode;
                        _cells.GetCell(2, k).Value = "" + woCell.GetCell(2, i).Value;
                        _cells.GetCell(2, k).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn06, k).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:GetDurationWOList()", ex.Message);
                        k++;
                    }
                    finally
                    {
                        _cells.GetCell(2, i).Select();
                        i++;
                    }
                }
            }
            else
            {
                MessageBox.Show(@"Debe seleccionar un Distrito en la hoja de " + SheetName01);
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ExecuteDurationWoActions()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName06, ResultColumn06);

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow06 + 1;
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            while (!string.IsNullOrWhiteSpace(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value)))
            {
                try
                {
                    var districtCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value);
                    var wo = WorkOrderActions.GetNewWorkOrderDto(_cells.GetEmptyIfNull(_cells.GetCell(2, i).Value));
                    var duration = new WorkOrderDuration
                    {
                        jobDurationsDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value),
                        jobDurationsCode = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value)),
                        jobDurationsStart = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                        jobDurationsFinish = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)
                    };
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value).ToUpper();
                    switch (action)
                    {
                        case "CREAR":
                            {
                                WorkOrderActions.CreateWorkOrderDuration(urlService, opSheet, districtCode, wo, duration);
                                _cells.GetCell(ResultColumn06, i).Value = "CREADO";
                                _cells.GetCell(ResultColumn06, i).Style = StyleConstants.Success;
                                _cells.GetCell(7, i).Value = "";//Para evitar duplicados por repetición
                            }
                            break;
                        case "ELIMINAR":
                            {
                                WorkOrderActions.DeleteWorkOrderDuration(urlService, opSheet, districtCode, wo, duration);
                                _cells.GetCell(ResultColumn06, i).Value = "ELIMINADO";
                                _cells.GetCell(ResultColumn06, i).Style = StyleConstants.Success;
                            }
                            break;
                        default:
                            _cells.GetCell(ResultColumn06, i).Value = "---";
                            break;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn06, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn06, i).Style = StyleConstants.Error;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteDurationWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn06, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReviewRefCodesList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowD04 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var district = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;
                    var workOrder = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);

                    var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, workOrder);
                    if (wo == null || wo.GetWorkOrderDto().no == null)
                        throw new Exception("WORK ORDER NO ENCONTRADA");

                    var woRefCodes = WorkOrderActions.GetWorkOrderReferenceCodes(_eFunctions, urlService, opSheet, district, workOrder);
                    //GENERAL
                    _cells.GetCell(3, i).Value = "'" + woRefCodes.WorkRequest;
                    _cells.GetCell(4, i).Value = "'" + woRefCodes.ComentariosDuraciones;
                    _cells.GetCell(5, i).Value = "'" + woRefCodes.ComentariosDuracionesText;
                    _cells.GetCell(6, i).Value = "'" + woRefCodes.EmpleadoId;
                    _cells.GetCell(7, i).Value = "'" + woRefCodes.NroComponente;
                    _cells.GetCell(8, i).Value = "'" + woRefCodes.P1EqLivMed;
                    _cells.GetCell(9, i).Value = "'" + woRefCodes.P2EqMovilMinero;
                    _cells.GetCell(10, i).Value = "'" + woRefCodes.P3ManejoSustPeligrosa;
                    _cells.GetCell(11, i).Value = "'" + woRefCodes.P4GuardasEquipo;
                    _cells.GetCell(12, i).Value = "'" + woRefCodes.P5Aislamiento;
                    _cells.GetCell(13, i).Value = "'" + woRefCodes.P6TrabajosAltura;
                    _cells.GetCell(14, i).Value = "'" + woRefCodes.P7ManejoCargas;
                    _cells.GetCell(15, i).Value = "'" + woRefCodes.ProyectoIcn;
                    _cells.GetCell(16, i).Value = "'" + woRefCodes.Reembolsable;
                    _cells.GetCell(17, i).Value = "'" + woRefCodes.FechaNoConforme;
                    _cells.GetCell(18, i).Value = "'" + woRefCodes.FechaNoConformeText;
                    _cells.GetCell(19, i).Value = "'" + woRefCodes.NoConforme;
                    _cells.GetCell(20, i).Value = "'" + woRefCodes.FechaEjecucion;
                    _cells.GetCell(21, i).Value = "'" + woRefCodes.HoraIngreso;
                    _cells.GetCell(22, i).Value = "'" + woRefCodes.HoraSalida;
                    _cells.GetCell(23, i).Value = "'" + woRefCodes.NombreBuque;
                    _cells.GetCell(24, i).Value = "'" + woRefCodes.CalificacionEncuesta;
                    _cells.GetCell(25, i).Value = "'" + woRefCodes.TareaCritica;
                    _cells.GetCell(26, i).Value = "'" + woRefCodes.Garantia;
                    _cells.GetCell(27, i).Value = "'" + woRefCodes.GarantiaText;
                    _cells.GetCell(28, i).Value = "'" + woRefCodes.CodigoCertificacion;
                    _cells.GetCell(29, i).Value = "'" + woRefCodes.FechaEntrega;
                    _cells.GetCell(30, i).Value = "'" + woRefCodes.RelacionarEv;
                    _cells.GetCell(31, i).Value = "'" + woRefCodes.Departamento;
                    _cells.GetCell(32, i).Value = "'" + woRefCodes.Localizacion;
                    _cells.GetCell(33, i).Value = "'" + woRefCodes.MetodoContacto;
                    _cells.GetCell(34, i).Value = "'" + woRefCodes.MetodoContactoText;

                    _cells.GetCell(ResultColumnD04, i).Value = "CONSULTADO";
                    _cells.GetCell(ResultColumnD04, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD04, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void UpdateReferenceCodes()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameD04, ResultColumnD04);

            var i = TitleRowD04 + 1;
            const int validationRow = TitleRowD04 - 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    //GENERAL
                    var district = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    var woRefCodes = new WorkOrderReferenceCodes
                    {
                        WorkRequest = MyUtilities.IsTrue(_cells.GetCell(03, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(03, i).Value) : null,
                        ComentariosDuraciones = MyUtilities.IsTrue(_cells.GetCell(04, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(04, i).Value) : null,
                        ComentariosDuracionesText = MyUtilities.IsTrue(_cells.GetCell(05, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(05, i).Value) : null,
                        EmpleadoId = MyUtilities.IsTrue(_cells.GetCell(06, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(06, i).Value) : null,
                        NroComponente = MyUtilities.IsTrue(_cells.GetCell(07, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(07, i).Value) : null,
                        P1EqLivMed = MyUtilities.IsTrue(_cells.GetCell(08, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(08, i).Value) : null,
                        P2EqMovilMinero = MyUtilities.IsTrue(_cells.GetCell(09, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(09, i).Value) : null,
                        P3ManejoSustPeligrosa = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null,
                        P4GuardasEquipo = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null,
                        P5Aislamiento = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null,
                        P6TrabajosAltura = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null,
                        P7ManejoCargas = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null,
                        ProyectoIcn = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null,
                        Reembolsable = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null,
                        FechaNoConforme = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null,
                        FechaNoConformeText = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null,
                        NoConforme = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null,
                        FechaEjecucion = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null,
                        HoraIngreso = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null,
                        HoraSalida = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null,
                        NombreBuque = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null,
                        CalificacionEncuesta = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null,
                        TareaCritica = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null,
                        Garantia = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null,
                        GarantiaText = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null,
                        CodigoCertificacion = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null,
                        FechaEntrega = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null,
                        RelacionarEv = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null,
                        Departamento = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null,
                        Localizacion = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null,
                        MetodoContacto = MyUtilities.IsTrue(_cells.GetCell(33, validationRow).Value) ? MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(33, i).Value)) : null,
                        MetodoContactoText = MyUtilities.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null
                    };


                    var replyRefCode = WorkOrderActions.UpdateWorkOrderReferenceCodes(_eFunctions, urlService, opSheet, district, workOrder, woRefCodes);

                    if (replyRefCode.Errors != null && replyRefCode.Errors.Length > 0)
                    {
                        var errorList = "";
                        // ReSharper disable once LoopCanBeConvertedToQuery
                        foreach (var error in replyRefCode.Errors)
                            errorList = errorList + "\nError: " + error;

                        _cells.GetCell(ResultColumnD04, i).Value = replyRefCode.Message + errorList;
                        _cells.GetCell(1, i).Style = StyleConstants.Warning;
                        _cells.GetCell(ResultColumnD04, i).Style = StyleConstants.Warning;
                    }
                    else
                    {
                        _cells.GetCell(ResultColumnD04, i).Value = "ACTUALIZADO";
                        _cells.GetCell(1, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumnD04, i).Style = StyleConstants.Success;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD04, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnD04, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateReferenceCodes()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumnD04, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnFormatCriticalControls_Click(object sender, RibbonControlEventArgs e)
        {
            FormatCriticalControls();
        }
        private void FormatCriticalControls()
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
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 4
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameCc01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "CONTROLES CRÍTICOS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");


                _cells.GetCell("A3").AddComment("--SUPERINTENDENCIA--\n" +
                    "FFCC, PBV, PTAS\n");
                _cells.GetCell("A3").Comment.Shape.TextFrame.AutoSize = true;

                var quartermasterList = new List<string> { "PUERTO BOLIVAR", "FERROCARRIL", "PLANTAS DE CARBON" };
                _cells.GetCell("A3").Value = SearchFieldCriteriaType.Quartermaster.Value;
                _cells.GetCell("A3").Style = StyleConstants.Option;
                _cells.GetCell("B3").Style = StyleConstants.Select;
                _cells.SetValidationList(_cells.GetCell("B3"), quartermasterList, ValidationSheetName, 1);

                _cells.GetRange(1, TitleRowCc01, ResultColumnCc01, TitleRowCc01).Style = StyleConstants.TitleRequired;

                //GENERAL

                _cells.GetCell(01, TitleRowCc01).Value = "Work Order";
                _cells.GetCell(02, TitleRowCc01).Value = "Tarea Nro";
                _cells.GetCell(03, TitleRowCc01).Value = "Descripción Tarea";
                _cells.GetCell(04, TitleRowCc01).Value = "Bow Tie";
                _cells.GetCell(05, TitleRowCc01).Value = "Descripción General";
                _cells.GetCell(06, TitleRowCc01).Value = "Equipo No";
                _cells.GetCell(07, TitleRowCc01).Value = "SuperIntendencia";
                _cells.GetCell(08, TitleRowCc01).Value = "Departamento";
                _cells.GetCell(09, TitleRowCc01).Value = "Asignado";
                _cells.GetCell(10, TitleRowCc01).Value = "Fecha Planeada";
                _cells.GetCell(11, TitleRowCc01).Value = "Fecha Creación";
                _cells.GetCell(12, TitleRowCc01).Value = "Maint Sch Task";
                _cells.GetCell(13, TitleRowCc01).Value = "Standard Job";
                _cells.GetCell(14, TitleRowCc01).Value = "Estado";
                _cells.GetCell(15, TitleRowCc01).Value = "Cód Completado";
                _cells.GetCell(16, TitleRowCc01).Value = "Completado Por";
                _cells.GetCell(17, TitleRowCc01).Value = "Fecha de Completado";
                _cells.GetCell(18, TitleRowCc01).Value = "Frecuencia";
                _cells.GetCell(19, TitleRowCc01).Value = "Texto Instrucciones";

                _cells.GetCell(ResultColumnCc01, TitleRowCc01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnCc01, TitleRowCc01).Style = StyleConstants.TitleResult;

                _cells.FormatAsTable(_cells.GetRange(1, TitleRowCc01, ResultColumnCc01, TitleRowCc01 + 1), TableNameCc01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatQuality()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void btnReviewCriticalControls_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameCc01))
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReviewCriticalControlsList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewCriticalControlsList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void ReviewCriticalControlsList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableNameCc01);

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();

            //Obtengo los valores de las opciones de búsqueda
            var district = "ICOR";
            var searchCriteriaKey1Text = SearchFieldCriteriaType.Quartermaster.Value;
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;

            var listcc = CriticalControlActions.FetchCriticalControl(_eFunctions, urlService, opSheet, district, searchCriteriaKey1, searchCriteriaValue1);
            var i = TitleRowCc01 + 1;
            foreach (var cc in listcc)
            {
                try
                {
                    //GENERAL
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumnCc01, i).Style = StyleConstants.Normal;
                    _cells.GetCell(1, i).Value = "'" + cc.WorkOrder;
                    _cells.GetCell(2, i).Value = "'" + cc.TaskNo;
                    _cells.GetCell(3, i).Value = "" + cc.TaskDescription;
                    _cells.GetCell(4, i).Value = "" + cc.CriticalDescription;
                    _cells.GetCell(5, i).Value = "" + cc.WorkOrderDescription;
                    _cells.GetCell(6, i).Value = "'" + cc.EquipmentNo;
                    _cells.GetCell(7, i).Value = "" + cc.Quartermaster;
                    _cells.GetCell(8, i).Value = "" + cc.Department;
                    _cells.GetCell(9, i).Value = "" + cc.AssignPerson;
                    _cells.GetCell(10, i).Value = "" + cc.PlanStartDate;
                    _cells.GetCell(11, i).Value = "" + cc.RaisedDate;
                    _cells.GetCell(12, i).Value = "'" + cc.MaintSchTask;
                    _cells.GetCell(13, i).Value = "'" + cc.StdJobNo;
                    _cells.GetCell(14, i).Value = "" + cc.Status;
                    _cells.GetCell(15, i).Value = "'" + cc.CompletedCode;
                    _cells.GetCell(16, i).Value = "" + cc.CompletedBy;
                    _cells.GetCell(17, i).Value = "" + cc.CompletedDate;
                    _cells.GetCell(18, i).Value = "" + cc.FrequencyText;
                    _cells.GetCell(19, i).Value = "" + cc.InstructionsText;
                    _cells.GetCell(19, i).WrapText = false;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnCc01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewCriticalControlsList()", ex.Message);
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

        private void btnReReviewCritialControls_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameCc01))
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ReReviewCriticalControlsList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReReviewCriticalControlsList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void ReReviewCriticalControlsList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //Obtengo los valores de las opciones de búsqueda
            var district = "ICOR";

            var i = TitleRowCc01 + 1;
            while (!string.IsNullOrWhiteSpace(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value)))
            {
                try
                {
                    var workOrder = "" + _cells.GetCell(1, i).Value;
                    var woTask = "" + _cells.GetCell(2, i).Value;
                    var cc = CriticalControlActions.FetchCriticalControl(_eFunctions, urlService, opSheet, district, workOrder, woTask);
                    //GENERAL
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumnCc01, i).Style = StyleConstants.Normal;
                    _cells.GetCell(1, i).Value = "'" + cc.WorkOrder;
                    _cells.GetCell(2, i).Value = "'" + cc.TaskNo;
                    _cells.GetCell(3, i).Value = "" + cc.TaskDescription;
                    _cells.GetCell(4, i).Value = "" + cc.CriticalDescription;
                    _cells.GetCell(5, i).Value = "" + cc.WorkOrderDescription;
                    _cells.GetCell(6, i).Value = "'" + cc.EquipmentNo;
                    _cells.GetCell(7, i).Value = "" + cc.Quartermaster;
                    _cells.GetCell(8, i).Value = "" + cc.Department;
                    _cells.GetCell(9, i).Value = "" + cc.AssignPerson;
                    _cells.GetCell(10, i).Value = "" + cc.PlanStartDate;
                    _cells.GetCell(11, i).Value = "" + cc.RaisedDate;
                    _cells.GetCell(12, i).Value = "'" + cc.MaintSchTask;
                    _cells.GetCell(13, i).Value = "'" + cc.StdJobNo;
                    _cells.GetCell(14, i).Value = "" + cc.Status;
                    _cells.GetCell(15, i).Value = "'" + cc.CompletedCode;
                    _cells.GetCell(16, i).Value = "" + cc.CompletedBy;
                    _cells.GetCell(17, i).Value = "" + cc.CompletedDate;
                    _cells.GetCell(18, i).Value = "" + cc.FrequencyText;
                    _cells.GetCell(19, i).Value = "" + cc.InstructionsText;
                    _cells.GetCell(19, i).WrapText = false;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnCc01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewCriticalControlsList()", ex.Message);
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

        private void btnExportCriticalControls_Click(object sender, RibbonControlEventArgs e)
        {
            ExportRichTextBox();
        }

        private void ExportRichTextBox()
        {
            var dc = new CriticalControlDefaultExport();

            //Creo el formulario de Opciones del Export
            #region Formulario Export
            //CheckBoxes
            var cbWorkOrder = new CheckBox { AutoSize = true, Text = @"Orden de Trabajo", Checked = dc.WorkOrder };
            var cbTaskNo = new CheckBox { AutoSize = true, Text = @"Tarea No", Checked = dc.TaskNo };
            var cbTaskDescription = new CheckBox { AutoSize = true, Text = @"Descripción Tarea", Checked = dc.TaskDescription };
            var cbWorkOrderDescription = new CheckBox { AutoSize = true, Text = @"Descripción General", Checked = dc.WorkOrderDescription };
            var cbCriticalDescription = new CheckBox { AutoSize = true, Text = @"Descripción BT", Checked = dc.CriticalDescription };
            var cbEquipmentNo = new CheckBox { AutoSize = true, Text = @"Equipo.", Checked = dc.EquipmentNo };
            var cbAssignPerson = new CheckBox { AutoSize = true, Text = @"Responsable", Checked = dc.AssignPerson };
            var cbDepartment = new CheckBox { AutoSize = true, Text = @"Departamento", Checked = dc.Department };
            var cbQuartermaster = new CheckBox { AutoSize = true, Text = @"SuperIntendencia", Checked = dc.Quartermaster };
            var cbPlanStartDate = new CheckBox { AutoSize = true, Text = @"Fecha Planeada", Checked = dc.PlanStartDate };
            var cbRaisedDate = new CheckBox { AutoSize = true, Text = @"Fecha Origen", Checked = dc.RaisedDate };
            var cbMaintSchTask = new CheckBox { AutoSize = true, Text = @"Mst No.", Checked = dc.MaintSchTask };
            var cbStdJobNo = new CheckBox { AutoSize = true, Text = @"Estándar Job No.", Checked = dc.StdJobNo };
            var cbStatus = new CheckBox { AutoSize = true, Text = @"Estado", Checked = dc.Status };
            var cbCompletedCode = new CheckBox { AutoSize = true, Text = @"Código Completado", Checked = dc.CompletedCode };
            var cbCompletedBy = new CheckBox { AutoSize = true, Text = @"Completado Por", Checked = dc.CompletedBy };
            var cbCompletedDate = new CheckBox { AutoSize = true, Text = @"Fecha Completado", Checked = dc.CompletedDate };
            var cbInstructionsText = new CheckBox { AutoSize = true, Text = @"Instrucciones", Checked = dc.InstructionsText };
            var cbFrequencyText = new CheckBox { AutoSize = true, Text = @"Frecuencia", Checked = dc.FrequencyText };

            var comboList = new List<CheckBox>
            {
                cbWorkOrder,
                cbTaskNo,
                cbTaskDescription,
                cbWorkOrderDescription,
                cbCriticalDescription,
                cbEquipmentNo,
                cbAssignPerson,
                cbDepartment,
                cbQuartermaster,
                cbPlanStartDate,
                cbRaisedDate,
                cbMaintSchTask,
                cbStdJobNo,
                cbStatus,
                cbCompletedCode,
                cbCompletedBy,
                cbCompletedDate,
                cbInstructionsText,
                cbFrequencyText
            };

            //Button Pane
            var okButton = new Button { DialogResult = DialogResult.OK, Text = @"Generar RTF" };
            var cancelButton = new Button { DialogResult = DialogResult.Cancel, Text = @"Cancelar" };
            var btnPane = new FlowLayoutPanel()
            {
                FlowDirection = FlowDirection.LeftToRight,

            };

            btnPane.Controls.Add(okButton);
            btnPane.Controls.Add(cancelButton);

            //Tab Pane
            var tabPane = new TableLayoutPanel { AutoSize = true, Padding = new Padding(9), ColumnCount = 2 };
            var comboArray = comboList.ToArray();
            for (var k = 0; k < comboArray.Length; k++)
                tabPane.Controls.Add(comboArray[k], 0, k);
            tabPane.Controls.Add(btnPane);
            tabPane.SetColumnSpan(btnPane, 2);
            var exportForm = new Form
            {
                Text = @"Exportar...",
                AutoSize = true
            };
            exportForm.Controls.Add(tabPane);
            #endregion
            if (exportForm.ShowDialog() != DialogResult.OK) return;
            _cells.SetCursorWait();

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);


            var i = TitleRowCc01 + 1;
            try
            {
                //Recorremos la hoja para general el listado a convertir a RTF
                var controlList = new List<CriticalControl>();
                while (!string.IsNullOrWhiteSpace(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value)))
                {
                    var cc = new CriticalControl
                    {
                        WorkOrder = "" + _cells.GetCell(1, i).Value,
                        TaskNo = "" + _cells.GetCell(2, i).Value,
                        TaskDescription = "" + _cells.GetCell(3, i).Value,
                        CriticalDescription = "" + _cells.GetCell(4, i).Value,
                        WorkOrderDescription = "" + _cells.GetCell(5, i).Value,
                        EquipmentNo = "" + _cells.GetCell(6, i).Value,
                        Quartermaster = "" + _cells.GetCell(7, i).Value,
                        Department = "" + _cells.GetCell(8, i).Value,
                        AssignPerson = "" + _cells.GetCell(9, i).Value,
                        PlanStartDate = "" + _cells.GetCell(10, i).Value,
                        RaisedDate = "" + _cells.GetCell(11, i).Value,
                        MaintSchTask = "" + _cells.GetCell(12, i).Value,
                        StdJobNo = "" + _cells.GetCell(13, i).Value,
                        Status = "" + _cells.GetCell(14, i).Value,
                        CompletedCode = "" + _cells.GetCell(15, i).Value,
                        CompletedBy = "" + _cells.GetCell(16, i).Value,
                        CompletedDate = "" + _cells.GetCell(17, i).Value,
                        FrequencyText = "" + _cells.GetCell(18, i).Value,
                        InstructionsText = "" + _cells.GetCell(19, i).Value
                    };

                    controlList.Add(cc);
                    i++;
                }

                #region Formulario RTF

                exportForm.Controls.Remove(tabPane);

                var exportOptions = new CriticalControlDefaultExport
                {
                    WorkOrder = cbWorkOrder.Checked,
                    TaskNo = cbTaskNo.Checked,
                    TaskDescription = cbTaskDescription.Checked,
                    WorkOrderDescription = cbWorkOrderDescription.Checked,
                    CriticalDescription = cbCriticalDescription.Checked,
                    EquipmentNo = cbEquipmentNo.Checked,
                    AssignPerson = cbAssignPerson.Checked,
                    Department = cbDepartment.Checked,
                    Quartermaster = cbQuartermaster.Checked,
                    PlanStartDate = cbPlanStartDate.Checked,
                    RaisedDate = cbRaisedDate.Checked,
                    MaintSchTask = cbMaintSchTask.Checked,
                    StdJobNo = cbStdJobNo.Checked,
                    Status = cbStatus.Checked,
                    CompletedCode = cbCompletedCode.Checked,
                    CompletedBy = cbCompletedBy.Checked,
                    CompletedDate = cbCompletedDate.Checked,
                    InstructionsText = cbInstructionsText.Checked,
                    FrequencyText = cbFrequencyText.Checked
                };
                var textBox = new RichTextBox
                {
                    Dock = DockStyle.Fill,
                    Rtf = CriticalControlActions.GetStringForExport(controlList, exportOptions)
                };
                exportForm.Controls.Add(textBox);
                exportForm.Controls.Add(cancelButton);
                exportForm.Width = 800;
                exportForm.Width = 600;
                #endregion
                exportForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"No se han podido generar los resultados. Se ha encontrado un error en la línea " + i);
                Debugger.LogError("RibbonEllipse.cs:ExportCriticalControlsList()", ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnReviewTasks_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                _thread = new Thread(ReviewWoTasks);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void ReviewWoTasks()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(TableName02);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var stOpContext = StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, true);

            var woCells = new ExcelStyleCells(_excelApp, SheetName01);
            woCells.SetAlwaysActiveSheet(false);

            var j = TitleRow01 + 1;//itera según cada estándar
            var i = TitleRow02 + 1;//itera la celda para cada tarea

            while (!string.IsNullOrEmpty("" + woCells.GetCell(3, j).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(woCells.GetCell(2, 3).Value2);
                    var workOrder = _cells.GetEmptyIfNull(woCells.GetCell(2, j).Value2);

                    var taskList = WorkOrderActions.FetchWorkOrderTask(_eFunctions, districtCode, workOrder, "");


                    foreach (var task in taskList)
                    {
                        //Para resetear el estilo
                        _cells.GetRange(1, i, ResultColumn02, i).Style = StyleConstants.Normal;
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + task.DistrictCode;
                        _cells.GetCell(2, i).Value = "" + task.WorkGroup;
                        _cells.GetCell(3, i).Value = "'" + task.WorkOrder;
                        _cells.GetCell(4, i).Value = "" + task.WorkOrderDescription;
                        //ACTION
                        _cells.GetCell(5, i).Value = "M";
                        //GENERAL
                        _cells.GetCell(6, i).Value = "'" + task.WoTaskNo;
                        _cells.GetCell(7, i).Value = "" + task.WoTaskDesc;
                        _cells.GetCell(8, i).Value = "'" + task.JobDescCode;
                        _cells.GetCell(9, i).Value = "'" + task.SafetyInstr;
                        _cells.GetCell(10, i).Value = "'" + task.CompleteInstr;
                        _cells.GetCell(11, i).Value = "'" + task.ComplTextCode;
                        //PLANNING
                        _cells.GetCell(12, i).Value = "" + task.AssignPerson;
                        _cells.GetCell(13, i).Value = "'" + task.EstimatedMachHrs;
                        _cells.GetCell(14, i).Value = "'" + task.PlanStartDate;

                        //RECURSOS
                        _cells.GetCell(15, i).Value = "" + task.EstimatedDurationsHrs;
                        _cells.GetCell(16, i).Value = "" + task.NoLabor;
                        _cells.GetCell(17, i).Value = "" + task.NoMaterial;
                        //APL
                        _cells.GetCell(18, i).Value = "'" + task.AplEquipmentGrpId;
                        _cells.GetCell(19, i).Value = "'" + task.AplType;
                        _cells.GetCell(20, i).Value = "'" + task.AplCompCode;
                        _cells.GetCell(21, i).Value = "'" + task.AplCompModCode;
                        _cells.GetCell(22, i).Value = "'" + task.AplSeqNo;
                        _cells.GetRange(17, i, 22, i).Style = !string.IsNullOrWhiteSpace(task.AplType)
                            ? StyleConstants.Error : StyleConstants.Normal;

                        var stdTextId = "WA" + task.DistrictCode + task.WorkOrder + task.WoTaskNo;
                        _cells.GetCell(23, i).Value = StdText.GetText(urlService, stOpContext, stdTextId);
                        _cells.GetCell(23, i).WrapText = false;
                        _cells.GetCell(3, i).Select();
                        i++;//aumenta tarea
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(1, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(1, j).Value2);
                    _cells.GetCell(2, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(2, j).Value2);
                    _cells.GetCell(3, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(3, j).Value2);
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn02, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:ReviewWoTasks()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta wo
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        private void btnReviewRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewRequirements);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void ReviewRequirements()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(TableName03);

            var taskCells = new ExcelStyleCells(_excelApp, SheetName02);
            taskCells.SetAlwaysActiveSheet(false);

            var j = TitleRow02 + 1;//itera según cada tarea
            var i = TitleRow03 + 1;//itera la celda para cada requerimiento

            var list = new List<TaskRequirement>();

            while (!string.IsNullOrEmpty("" + taskCells.GetCell(3, j).Value) && !string.IsNullOrEmpty("" + taskCells.GetCell(6, j).Value))
            {
                list.Add(new TaskRequirement
                {
                    DistrictCode = _cells.GetEmptyIfNull(taskCells.GetCell(1, j).Value2),
                    WorkGroup = _cells.GetEmptyIfNull(taskCells.GetCell(2, j).Value2),
                    WorkOrder = _cells.GetEmptyIfNull(taskCells.GetCell(3, j).Value2),
                    WoTaskNo = _cells.GetEmptyIfNull(taskCells.GetCell(6, j).Value2)
                });
                j++;
            }

            var distinctItems = list.GroupBy(x => new { x.DistrictCode, x.WorkGroup, x.WorkOrder, x.WoTaskNo }).Select(y => y.First());

            foreach (var d in distinctItems)
            {
                try
                {
                    var reqList = WorkOrderActions.FetchTaskRequirements(_eFunctions, d.DistrictCode, d.WorkGroup, d.WorkOrder, d.WoTaskNo);

                    var distinctReqList = reqList.GroupBy(x => new { x.DistrictCode, x.WorkGroup, x.WorkOrder, x.WoTaskNo, x.ReqCode}).Select(y => y.First());

                    foreach (var req in distinctReqList)
                    {
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + req.DistrictCode; //DistrictCode
                        _cells.GetCell(2, i).Value = "" + req.WorkGroup;    //WorkGroup
                        _cells.GetCell(3, i).Value = "" + req.WorkOrder;    //WorkOrder 
                        _cells.GetCell(4, i).Value = "" + req.WoTaskNo;     //WoTaskNo 
                        _cells.GetCell(5, i).Value = "" + req.WoTaskDesc;   //WoTaskDesc 
                        _cells.GetCell(6, i).Value = "M";
                        _cells.GetCell(7, i).Value = "" + req.ReqType;      //ReqType 
                        _cells.GetCell(8, i).Value = "" + req.SeqNo;        //SeqNo 
                        _cells.GetCell(9, i).Value = "" + req.ReqCode;      //ReqCode
                        _cells.GetCell(10, i).Value = "" + req.ReqDesc;     //ReqDesc
                        _cells.GetCell(11, i).Value = "" + req.UoM;         //UoM
                        _cells.GetCell(12, i).Value = "" + req.QtyReq;      //QtyReq
                        _cells.GetCell(13, i).Value = "" + req.QtyIss;      //QtyIss
                        _cells.GetCell(14, i).Value = "" + req.HrsReq;      //HrsReq
                        _cells.GetCell(15, i).Value = "" + req.HrsReal;     //HrsReal
                        _cells.GetCell(ResultColumn03, i).Select();
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(1, i).Value = d.DistrictCode;
                    _cells.GetCell(2, i).Value = d.WorkGroup;
                    _cells.GetCell(3, i).Value = d.WorkOrder;
                    _cells.GetCell(4, i).Value = d.WoTaskNo;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn03, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:ReviewRequirements()", ex.Message);
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnExecuteTaskActions_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(ExecuteTaskActions);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        /// <summary>
        /// Ejecuta las acciones de tarea mediante el servicio EWS
        /// </summary>
        // ReSharper disable once UnusedMember.Local
        private void ExecuteTaskActions()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);
            var i = TitleRow02 + 1;

            var opSheet = new WorkOrderTaskService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);


            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                    var woTask = new WorkOrderTask
                    {
                        DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        WoTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),
                        WoTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        JobDescCode = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        SafetyInstr = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value),
                        CompleteInstr = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                        ComplTextCode = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value),
                        AssignPerson = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value),
                        EstimatedMachHrs = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value),
                        PlanStartDate = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value),
                        EstimatedDurationsHrs = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value),
                        NoLabor = _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value),
                        NoMaterial = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value),
                        AplEquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value),
                        AplType = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value),
                        AplCompCode = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value),
                        AplCompModCode = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value),
                        AplSeqNo = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value),
                        ExtTaskText = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value)
                    };

                    woTask.SetWorkOrderDto(woTask.WorkOrder);

                    if (string.IsNullOrWhiteSpace(action))
                        continue;

                    if (action.Equals("M"))
                    {
                        WorkOrderActions.ModifyWorkOrderTask(urlService, opSheet, woTask, true);
                        WorkOrderActions.SetWorkOrderTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, woTask);
                    }
                    else if (action.Equals("C"))
                    {
                        WorkOrderActions.CreateWorkOrderTask(urlService, opSheet, woTask, true);
                        WorkOrderActions.SetWorkOrderTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, woTask);
                    }
                    else if (action.Equals("D"))
                    {
                        WorkOrderActions.DeleteWorkOrderTask(urlService, opSheet, woTask, true);
                        WorkOrderActions.SetWorkOrderTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, woTask);
                    }
                    else
                        continue;

                    _cells.GetCell(ResultColumn02, i).Value = "OK";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteTaskActions()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn02, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnExecuteRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                _thread = new Thread(ExecuteRequirementActions);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void ExecuteRequirementActions()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName03, ResultColumn03);
            var i = TitleRow03 + 1;

            var opSheetResource = new ResourceReqmntsService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };
            var opSheetMaterial = new MaterialReqmntsService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };
            var opSheetEquipment = new EquipmentReqmntsService.OperationContext()
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value) && !string.IsNullOrEmpty("" + _cells.GetCell(4, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var taskReq = new TaskRequirement();
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value);                         //_cells.GetCell(6, i).Value = "M";

                    taskReq.DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);                  //_cells.GetCell(1, i).Value = "" + req.DistrictCode; 
                    taskReq.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);                     //_cells.GetCell(2, i).Value = "" + req.WorkGroup;    
                    taskReq.WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);                     //_cells.GetCell(3, i).Value = "" + req.WorkOrder;     
                    taskReq.WoTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);                      //_cells.GetCell(4, i).Value = "" + req.WoTaskNo;      
                    taskReq.WoTaskNo = string.IsNullOrWhiteSpace(taskReq.WoTaskNo) ? "001" : taskReq.WoTaskNo;
                    taskReq.WoTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);                    //_cells.GetCell(5, i).Value = "" + req.WoTaskDesc;
                    taskReq.ReqType = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value);                       //_cells.GetCell(7, i).Value = "" + req.ReqType;       
                    taskReq.SeqNo = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value);                         //_cells.GetCell(8, i).Value = "" + req.SeqNo;         
                    taskReq.ReqCode = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value);                       //_cells.GetCell(9, i).Value = "" + req.ReqCode;      
                    taskReq.ReqDesc = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value);                      //_cells.GetCell(10, i).Value = "" + req.ReqDesc;
                    taskReq.UoM = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value);                          //_cells.GetCell(11, i).Value = "" + req.UoM;
                    taskReq.QtyReq = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value);                       //_cells.GetCell(12, i).Value = "" + req.QtyReq;       
                    taskReq.QtyIss = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value);                       //_cells.GetCell(13, i).Value = "" + req.QtyIss;      
                    taskReq.HrsReq = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value);                       //_cells.GetCell(14, i).Value = "" + req.HrsReq;      
                    taskReq.HrsReal = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value);                      //_cells.GetCell(15, i).Value = "" + req.HrsReal;     


                    if (string.IsNullOrWhiteSpace(action))
                        continue;
                    else if (action.Equals("C"))
                    {
                        if (taskReq.ReqType.Equals("LAB"))
                            WorkOrderActions.CreateTaskResource(urlService, opSheetResource, taskReq);
                        else if (taskReq.ReqType.Equals("MAT"))
                            WorkOrderActions.CreateTaskMaterial(urlService, opSheetMaterial, taskReq);
                        else if (taskReq.ReqType.Equals("EQU"))
                            WorkOrderActions.CreateTaskEquipment(urlService, opSheetEquipment, taskReq);
                    }
                    else if (action.Equals("M"))
                    {
                        if (taskReq.ReqType.Equals("LAB"))
                            WorkOrderActions.ModifyTaskResource(urlService, opSheetResource, taskReq);
                        else if (taskReq.ReqType.Equals("MAT"))
                            WorkOrderActions.ModifyTaskMaterial(urlService, opSheetMaterial, taskReq);
                        else if (taskReq.ReqType.Equals("EQU"))
                            WorkOrderActions.ModifyTaskEquipment(urlService, opSheetEquipment, taskReq);
                    }
                    else if (action.Equals("D"))
                    {
                        if (taskReq.ReqType.Equals("LAB"))
                            WorkOrderActions.DeleteTaskResource(urlService, opSheetResource, taskReq);
                        else if (taskReq.ReqType.Equals("MAT"))
                            WorkOrderActions.DeleteTaskMaterial(urlService, opSheetMaterial, taskReq);
                        else if (taskReq.ReqType.Equals("EQU"))
                            WorkOrderActions.DeleteTaskEquipment(urlService, opSheetEquipment, taskReq);
                    }
                    _cells.GetCell(ResultColumn03, i).Value = "OK";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteRequirementActions()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn03, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
    }

}
