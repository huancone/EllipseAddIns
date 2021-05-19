using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Utilities;
using EllipseWorkOrdersClassLibrary;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using SharedClassLibrary.Ellipse.Forms;
using EllipseStdTextClassLibrary;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse.Connections;
using WorkOrderTaskService = EllipseWorkOrdersClassLibrary.WorkOrderTaskService;
using WorkOrderService = EllipseWorkOrdersClassLibrary.WorkOrderService;
using ResourceReqmntsService = EllipseWorkOrdersClassLibrary.ResourceReqmntsService;
using MaterialReqmntsService = EllipseWorkOrdersClassLibrary.MaterialReqmntsService;
using EquipmentReqmntsService = EllipseWorkOrdersClassLibrary.EquipmentReqmntsService;

namespace EllipseWorkOrderExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const string SheetName01 = "WorkOrders";
        private const string SheetName02 = "Tasks";
        private const string SheetName03 = "Requirement";
        private const string SheetName04 = "CloseWorkOrders";
        private const string SheetName05 = "CloseCommentsWorkOrders";
        private const string SheetName06 = "DurationWorkOrders";
        private const string SheetName07 = "ProgressWorkOrders";
        private const string SheetName08 = "ToDoWorkOrders";

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
        private const int TitleRow08 = 5;

        private const int TitleRowD01 = 9;
        private const int TitleRowD02 = 6;
        private const int TitleRowD03 = 6;
        private const int TitleRowD04 = 6;
        private const int TitleRowQ01 = 7;
        private const int TitleRowCc01 = 6;

        private const int ResultColumn01 = 54;
        private const int ResultColumn02 = 31;
        private const int ResultColumn03 = 16;
        private const int ResultColumn04 = 8;
        private const int ResultColumn05 = 5;
        private const int ResultColumn06 = 9;
        private const int ResultColumn07 = 6;
        private const int ResultColumn08 = 13;

        private const int ResultColumnD01 = 56;
        private const int ResultColumnD02 = 8;
        private const int ResultColumnD03 = 3;
        private const int ResultColumnD04 = 38;
        private const int ResultColumnQ01 = 36;
        private const int ResultColumnCc01 = 20;

        private const string TableName01 = "WorkOrderTable";
        private const string TableName02 = "TaskTable";
        private const string TableName03 = "RequirementTable";
        private const string TableName04 = "WorkOrderCloseTable";
        private const string TableName05 = "WorkOrderCompleteTextTable";
        private const string TableName06 = "WorkOrderDurationTable";
        private const string TableName07 = "WorkOrderProgressTable";
        private const string TableName08 = "WorkOrderToDoTable";

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
            LoadSettings();
        }
        public void LoadSettings()
        {
            try
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

                settings.SetDefaultCustomSettingValue("FlagEstDuration", "Y");
                settings.SetDefaultCustomSettingValue("ValidateTaskPlanDates", "Y");
                settings.SetDefaultCustomSettingValue("IgnoreClosedStatus", "N");



                //Setting of Configuration Options from Config File (or default)
                try
                {
                    settings.LoadCustomSettings();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                var flagEstDur = MyUtilities.IsTrue(settings.GetCustomSettingValue("FlagEstDuration"));
                var valdTaskPlanDates = MyUtilities.IsTrue(settings.GetCustomSettingValue("ValidateTaskPlanDates"));
                var ignoreCldStat = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreClosedStatus"));

                cbFlagEstDuration.Checked = flagEstDur;
                cbValidateTaskPlanDates.Checked = valdTaskPlanDates;
                cbIgnoreClosedStatus.Checked = ignoreCldStat;
                //
                settings.SaveCustomSettings();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
 #region -- buttonActions -- 
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
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReReviewWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(CreateWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateWoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameD01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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

        private void btnCleanWorkOrderSheet_Click(object sender, RibbonControlEventArgs e)
        {
            CleanTable(TableName01);
            CleanTable(TableNameD01);
            CleanTable(TableNameD02);
            CleanTable(TableNameD03);
            CleanTable(TableNameD04);
        }

        private void btnCleanCloseSheets_Click(object sender, RibbonControlEventArgs e)
        {
            CleanTable(TableName04);
            CleanTable(TableName05);
        }
        private void btnCleanDuration_Click(object sender, RibbonControlEventArgs e)
        {
            CleanTable(TableName06);
        }
        private void btnReviewReferenceCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameD04))
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
            CleanTable(TableNameQ01);
        }
        private void btnReviewQuality_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetNameQ01))
                {
                    //si ya hay un thread corriendo que no se ha detenido
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
                    //si ya hay un thread corriendo que no se ha detenido
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

        #endregion
        private void ReviewWorkProgress()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
        
        
        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 4
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 5
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 6
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 7
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 8

                #region CONSTRUYO LA HOJA 1
                var titleRow = TitleRow01;
                var resultColumn = ResultColumn01;
                var tableName = TableName01;
                var sheetName = SheetName01;

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;
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

                _cells.GetRange(1, titleRow, resultColumn, titleRow).Style = StyleConstants.TitleOptional;
                for (var i = 4; i < resultColumn - 4; i++)
                {
                    _cells.GetCell(i, titleRow - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, titleRow - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, titleRow - 1).Value = "true";
                }

                //GENERAL
                _cells.GetCell(1, titleRow).Value = "WORK_GROUP";
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, titleRow).Value = "WORK_ORDER";
                _cells.GetCell(2, titleRow).AddComment("Ingrese solo el prefijo si quiere crear una orden con prefijo");
                _cells.GetCell(3, titleRow).Value = "WO_STATUS";
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleInformation;

                _cells.GetCell(4, titleRow - 2).Value = "GENERAL";
                _cells.GetCell(4, titleRow).Value = "DESCRIPTION";
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, titleRow).Value = "EQUIPMENT";
                _cells.GetCell(5, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, titleRow).Value = "COMP_CODE";
                _cells.GetCell(7, titleRow).Value = "MOD_CODE";

                var priorityCodes = MyUtilities.GetCodeList(WoTypeMtType.GetPriorityCodeList());
                var woTypeCodes = MyUtilities.GetCodeList(WoTypeMtType.GetWoTypeList());
                var mtTypeCodes = MyUtilities.GetCodeList(WoTypeMtType.GetMtTypeList());
                
                var usTypeCodes = new List<string>();
                foreach(var item in WorkOrderActions.GetUserStatusCodeList(_eFunctions))
                    usTypeCodes.Add(item.Code + " - " + item.Description);

                _cells.GetCell(8, titleRow).Value = "WO_TYPE";
                _cells.GetCell(8, titleRow).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(8, titleRow + 1), woTypeCodes, ValidationSheetName, 6, false);
                _cells.GetCell(9, titleRow).Value = "MT_TYPE";
                _cells.GetCell(9, titleRow).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(9, titleRow + 1), mtTypeCodes, ValidationSheetName, 7, false);
                _cells.GetCell(10, titleRow).Value = "WO_USER_STATUS";
                _cells.SetValidationList(_cells.GetCell(10, titleRow + 1), usTypeCodes, ValidationSheetName, 8, false);
                _cells.GetCell(11, titleRow).Value = "RAISED_DATE";
                _cells.GetCell(11, titleRow).AddComment("yyyyMMdd");
                _cells.GetCell(12, titleRow).Value = "RAISED_TIME";
                _cells.GetCell(12, titleRow).AddComment("hhmmss");
                _cells.GetCell(13, titleRow).Value = "ORIGINATOR_ID";
                _cells.GetCell(14, titleRow).Value = "ORIG_PRIORITY";
                _cells.GetCell(14, titleRow).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(14, titleRow + 1), priorityCodes, ValidationSheetName, 9, false);
                _cells.GetCell(15, titleRow).Value = "ORIG_DOC_TYPE";
                _cells.GetCell(16, titleRow).Value = "ORIG_DOC_NO";
                _cells.GetCell(17, titleRow).Value = "RELATED_WO";
                _cells.GetCell(18, titleRow).Value = "WORKREQUEST";
                _cells.GetCell(19, titleRow).Value = "STD_JOB";
                _cells.GetCell(20, titleRow).Value = "MST";
                _cells.GetCell(20, titleRow).Style = StyleConstants.TitleInformation;
                _cells.GetCell(20, titleRow - 1).Value = "N/A";

                _cells.GetRange(4, titleRow - 2, 20, titleRow - 2).Style = StyleConstants.Select;
                _cells.GetRange(4, titleRow - 2, 20, titleRow - 2).Merge();

                //PLANNING
                _cells.GetCell(21, titleRow - 2).Value = "PLANNING";
                _cells.GetCell(21, titleRow).Value = "AUTO_REQ";
                _cells.GetCell(21, titleRow).AddComment("Y/N");
                _cells.GetCell(22, titleRow).Value = "ASSIGN";
                _cells.GetCell(23, titleRow).Value = "PLAN_PRIORITY";
                _cells.GetCell(23, titleRow).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(23, titleRow + 1), ValidationSheetName, 9, false);
                _cells.GetCell(24, titleRow).Value = "REQ_START_DATE";
                _cells.GetCell(24, titleRow).AddComment("yyyyMMdd");
                _cells.GetCell(25, titleRow).Value = "REQ_START_TIME";
                _cells.GetCell(25, titleRow).AddComment("hhmmss");
                _cells.GetCell(26, titleRow).Value = "REQ_BY_DATE";
                _cells.GetCell(26, titleRow).AddComment("yyyyMMdd");
                _cells.GetCell(27, titleRow).Value = "REQ_BY_TIME";
                _cells.GetCell(27, titleRow).AddComment("hhmmss");
                _cells.GetCell(28, titleRow).Value = "PLAN_STR_DATE";
                _cells.GetCell(28, titleRow).AddComment("yyyyMMdd - Las fechas de plan solo se modificarán si el usuario tiene permisos de planeación/programación");
                _cells.GetCell(29, titleRow).Value = "PLAN_STR_TIME";
                _cells.GetCell(29, titleRow).AddComment("hhmmss");
                _cells.GetCell(30, titleRow).Value = "PLAN_FIN_DATE";
                _cells.GetCell(30, titleRow).AddComment("yyyyMMdd - El comportamiento de este campo depende de la tarea de la orden");
                _cells.GetCell(31, titleRow).Value = "PLAN_FIN_TIME";
                _cells.GetCell(31, titleRow).AddComment("hhmmss");
                _cells.GetCell(32, titleRow).Value = "UNIT_OF_WORK";
                _cells.GetCell(33, titleRow).Value = "UNITS_REQUIRED";
                _cells.GetCell(34, titleRow).Value = "PC/UNITS COMP";
                _cells.GetCell(34, titleRow - 1).Value = "N/A";
                _cells.GetCell(34, titleRow).Style = StyleConstants.TitleInformation;
                _cells.GetRange(21, titleRow - 2, 34, titleRow - 2).Style = StyleConstants.Select;
                _cells.GetRange(21, titleRow - 2, 34, titleRow - 2).Merge();

                //COST
                _cells.GetCell(35, titleRow - 2).Value = "COST";
                _cells.GetCell(35, titleRow).Value = "ACCOUNT_CODE";
                _cells.GetCell(36, titleRow).Value = "PROJECT_NO";
                _cells.GetCell(37, titleRow).Value = "PARENT_WO";
                _cells.GetRange(35, titleRow - 2, 37, titleRow - 2).Style = StyleConstants.Select;
                _cells.GetRange(35, titleRow - 2, 37, titleRow - 2).Merge();

                //JOB_CODES
                _cells.GetCell(38, titleRow - 2).Value = "JOB CODES/FALLAS";
                _cells.GetCell(38, titleRow - 2).AddComment("Debe seleccionar por lo menos un Job Code para las órdenes correctivas/reparación");
                _cells.GetCell(38, titleRow).Value = "JOBCODE_01";
                _cells.GetCell(39, titleRow).Value = "JOBCODE_02";
                _cells.GetCell(40, titleRow).Value = "JOBCODE_03";
                _cells.GetCell(41, titleRow).Value = "JOBCODE_04";
                _cells.GetCell(42, titleRow).Value = "JOBCODE_05";
                _cells.GetCell(43, titleRow).Value = "JOBCODE_06";
                _cells.GetCell(44, titleRow).Value = "JOBCODE_07";
                _cells.GetCell(45, titleRow).Value = "JOBCODE_08";
                _cells.GetCell(46, titleRow).Value = "JOBCODE_09";
                _cells.GetCell(47, titleRow).Value = "JOBCODE_10";
                _cells.GetCell(48, titleRow).Value = "LOCATION FR";
                _cells.GetCell(49, titleRow).Value = "PART FAILURE";
                _cells.GetRange(38, titleRow - 2, 49, titleRow - 2).Style = StyleConstants.Select;
                _cells.GetRange(38, titleRow - 2, 49, titleRow - 2).Merge();
                //COMPLETION INFO
                _cells.GetCell(50, titleRow - 2).Value = "COMPL.INFO";
                _cells.GetCell(50, titleRow).Value = "COMPL_COD";
                _cells.GetCell(50, titleRow).AddComment("Código de cierre de la orden");
                _cells.GetCell(51, titleRow).Value = "COMP_COMM";
                _cells.GetCell(51, titleRow).AddComment("Indica si una orden tiene comentarios de cierre");
                _cells.GetCell(52, titleRow).Value = "CLOSED DATE";
                _cells.GetCell(53, titleRow).Value = "COMPL_BY";
                _cells.GetRange(50, titleRow - 2, 53, titleRow - 2).Style = StyleConstants.Option;
                _cells.GetRange(50, titleRow, 53, titleRow).Style = StyleConstants.TitleInformation;
                _cells.GetRange(50, titleRow - 2, 53, titleRow - 1).Merge();

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 2 - WO TASKS
                titleRow = TitleRow02;
                resultColumn = ResultColumn02;
                tableName = TableName02;
                sheetName = SheetName02;

                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

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

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, titleRow - 1).Value = "WORK ORDER";
                _cells.GetRange(1, titleRow - 1, 5, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, titleRow - 1, 5, titleRow - 1).Merge();

                _cells.GetCell(1, titleRow).Value = "DISTRICT";
                _cells.GetCell(2, titleRow).Value = "WORK_GROUP";
                _cells.GetCell(3, titleRow).Value = "WORK_ORDER";
                _cells.GetCell(4, titleRow).Value = "WO_DESC";
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleInformation;
                //ACTION
                _cells.GetCell(5, titleRow).Value = "ACTION";
                _cells.GetCell(5, titleRow).Style = StyleConstants.TitleAction;

                var actionList = WorkOrderTaskActions.GetActionsList();
                var actionListComment = "";
                foreach (var item in actionList)
                    actionListComment += item + "\n";
                _cells.GetCell(5, titleRow).AddComment(actionListComment);
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), actionList);
                //GENERAL
                _cells.GetCell(6, titleRow - 1).Value = "GENERAL";
                _cells.GetRange(6, titleRow - 1, 11, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(6, titleRow - 1, 11, titleRow - 1).Merge();

                _cells.GetCell(6, titleRow).Value = "TASK_NO";
                _cells.GetCell(7, titleRow).Value = "WO_TASK_DESC";
                _cells.GetCell(8, titleRow).Value = "JOB_DESC_CODE";
                _cells.GetCell(9, titleRow).Value = "SAFETY_INST";
                _cells.GetCell(10, titleRow).Value = "COMPL_INST";
                _cells.GetCell(11, titleRow).Value = "COMPL_TEXT_CODE";

                //PLANNING
                _cells.GetCell(12, titleRow - 1).Value = "PLANNING";
                _cells.GetRange(12, titleRow - 1, 17, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(12, titleRow - 1, 17, titleRow - 1).Merge();

                _cells.GetCell(12, titleRow).Value = "ASSIGN_PERSON";
                _cells.GetCell(13, titleRow).Value = "EST_MACH_HRS";
                _cells.GetCell(14, titleRow).Value = "PLAN START DATE";
                _cells.GetCell(15, titleRow).Value = "PLAN START TIME";
                _cells.GetCell(16, titleRow).Value = "PLAN FINISH DATE";
                _cells.GetCell(17, titleRow).Value = "PLAN FINISH TIME";
                _cells.GetRange(12, titleRow, 17, titleRow).Style = StyleConstants.TitleOptional;

                //RECURSOS
                _cells.GetCell(18, titleRow - 1).Value = "RECURSOS";
                _cells.GetRange(18, titleRow - 1, 20, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(18, titleRow - 1, 20, titleRow - 1).Merge();

                _cells.GetCell(18, titleRow).Value = "EST_DUR_HRS";
                _cells.GetCell(18, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(19, titleRow).Value = "LABOR";
                _cells.GetCell(20, titleRow).Value = "MATERIAL";
                _cells.GetRange(18, titleRow, 20, titleRow).Style = StyleConstants.TitleInformation;

                //APL
                _cells.GetCell(21, titleRow - 1).Value = "APL";
                _cells.GetRange(21, titleRow - 1, 25, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(21, titleRow - 1, 25, titleRow - 1).Merge();

                _cells.GetCell(21, titleRow).Value = "EQUIP_GRP_ID";
                _cells.GetCell(22, titleRow).Value = "APL_TYPE";
                _cells.GetCell(23, titleRow).Value = "COMP_CODE";
                _cells.GetCell(24, titleRow).Value = "COMP_MOD_CODE";
                _cells.GetCell(25, titleRow).Value = "APL_SEQ_NO";

                _cells.GetRange(21, titleRow, 25, titleRow).Style = StyleConstants.TitleOptional;

                _cells.GetCell(26, titleRow - 1).Value = "DESCRIPTION";
                _cells.GetRange(26, titleRow - 1, 26, titleRow - 1).Style = StyleConstants.Option;
                //_cells.GetRange(26, titleRow - 1, 30, titleRow - 1).Merge();
                _cells.GetCell(26, titleRow).Value = "DESCRIPCION EXTENDIDA";
                _cells.GetCell(26, titleRow).Style = StyleConstants.TitleOptional;
                //CIERRE DE TAREA
                _cells.GetCell(27, titleRow - 1).Value = "COMPLETION";
                _cells.GetRange(27, titleRow - 1, 30, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(27, titleRow - 1, 30, titleRow - 1).Merge();

                var completeCodeList = _eFunctions.GetItemCodesString("SC");
                _cells.SetValidationList(_cells.GetCell(27, titleRow + 1), completeCodeList, ValidationSheetName, 10, false);
                _cells.GetCell(27, titleRow).Value = "CÓD. DE CIERRE";
                _cells.GetCell(27, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(28, titleRow).Value = "COMPLETADO POR";
                _cells.GetCell(28, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(29, titleRow).Value = "FECHA DE CIERRE";
                _cells.GetCell(29, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(30, titleRow).Value = "COMENTARIOS DE CIERRE";
                _cells.GetCell(30, titleRow).Style = StyleConstants.TitleOptional;
                //RESULTADO
                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 3 - WO TASK REQUIREMENTS
                titleRow = TitleRow03;
                resultColumn = ResultColumn03;
                tableName = TableName03;
                sheetName = SheetName03;
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

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

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, titleRow - 1).Value = "ORDEN DE TRABAJO/ TAREA";
                _cells.GetRange(1, titleRow - 1, 6, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, titleRow - 1, 6, titleRow - 1).Merge();

                _cells.GetCell(1, titleRow).Value = "Distrito";       //_cells.GetCell(1, i).Value = req.DistrictCode; 
                _cells.GetCell(2, titleRow).Value = "Grupo/UAS";     //_cells.GetCell(2, i).Value = req.WorkGroup;    
                _cells.GetCell(3, titleRow).Value = "OrdenTrabajo";          //_cells.GetCell(3, i).Value = req.WorkOrder;    
                _cells.GetCell(4, titleRow).Value = "Tarea No.";        //_cells.GetCell(4, i).Value = req.WoTaskNo;     
                _cells.GetCell(5, titleRow).Value = "Desc. Tarea";   //_cells.GetCell(5, i).Value = req.WoTaskDesc;    

                //ACTION
                _cells.GetCell(6, titleRow).Value = "Acción";
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleAction;
                _cells.GetCell(6, titleRow).AddComment("C: Crear Requerimiento \nM: Modificar Requerimiento \nD: Eliminar Requerimiento");
                _cells.SetValidationList(_cells.GetCell(6, titleRow + 1), new List<string> { "C", "M", "D" });
                //GENERAL
                _cells.GetCell(7, titleRow - 1).Value = "INFORMACIÓN DEL RECURSO";
                _cells.GetRange(7, titleRow - 1, 15, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetRange(7, titleRow - 1, 15, titleRow - 1).Merge();

                _cells.GetCell(7, titleRow).Value = "Tipo Recurso";       //_cells.GetCell(7, i).Value = "" + req.ReqType;
                _cells.GetCell(7, titleRow).AddComment("LAB: Labor\nMAT: Material\nEQP: Equipos");
                _cells.SetValidationList(_cells.GetCell(7, titleRow + 1), new List<string> { RequirementType.Labour.Key, RequirementType.Material.Key, RequirementType.Equipment.Key});


                _cells.GetCell(8, titleRow).Value = "Seq. No.";         //_cells.GetCell(8, i).Value = req.SeqNo;    
                _cells.GetCell(8, titleRow).AddComment("Aplica solo para Creación y Modificación de Requerimientos");
                _cells.GetCell(9, titleRow).Value = "Req.Code/StockCode\nEquip.Type";       //_cells.GetCell(9, i).Value = req.ReqCode;  
                _cells.GetCell(9, titleRow).AddComment("Recurso: Class+Code (Ver hoja de recursos) \nMaterial: StockCode\nEquipos: Equipment Type");
                _cells.GetCell(10, titleRow).Value = "Desc. Recurso";   //_cells.GetCell(10, i).Value = req.ReqDesc; 
                _cells.GetCell(11, titleRow).Value = "UoM";           //_cells.GetCell(11, i).Value = req.UoM;  
                _cells.GetCell(12, titleRow).Value = "Tamaño Estimado";       //_cells.GetCell(11, i).Value = req.QtyReq;  
                _cells.GetCell(12, titleRow).AddComment("Labor: Tamaño de Personal \nEquipo: Tamaño de Flota\nMateriales: N/A Siempre será 1");
                _cells.GetCell(13, titleRow).Value = "Un. Est.";       //_cells.GetCell(12, i).Value = req.QtyIss;  
                _cells.GetCell(13, titleRow).AddComment("Unidades Estimadas del recurso");
                _cells.GetCell(14, titleRow).Value = "Und. Real";      //_cells.GetCell(14, i).Value = req.HrsReal; 
                _cells.GetCell(14, titleRow).Style = StyleConstants.TitleInformation;
                _cells.GetCell(14, titleRow).AddComment("Unidades Reales del recurso. Los recursos reales de materiales y equipos son compartidos entre las diferentes tareas de la orden que tengan el mismo stock code. Se resaltarán cuando esto ocurra");
                _cells.GetCell(15, titleRow).Value = "Compartido";      //_cells.GetCell(14, i).Value = req.HrsReal; 
                _cells.GetCell(15, titleRow).Style = StyleConstants.TitleInformation;
                _cells.GetCell(15, titleRow).AddComment("Entre cuántas tareas se comparte este recurso. Si el valor es 0 es porque no es un recurso estimado");

                //RESULTADO
                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, titleRow + 1, resultColumn - 2, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion 

                #region CONSTRUYO LA HOJA 4 - CLOSE WO
                titleRow = TitleRow04;
                resultColumn = ResultColumn04;
                tableName = TableName04;
                sheetName = SheetName04;
                _excelApp.ActiveWorkbook.Sheets[4].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

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
                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, titleRow).Value = "WORK_ORDER";
                _cells.GetCell(2, titleRow).Value = "CLOSED_DATE";
                _cells.GetCell(2, titleRow).AddComment("yyyyMMdd");
                _cells.GetCell(3, titleRow).Value = "CLOSED_TIME";
                _cells.GetCell(3, titleRow).AddComment("hhmmss");
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, titleRow).Value = "COMPLETED_BY";
                _cells.GetCell(5, titleRow).Value = "COMPLETED_CODE";
                _cells.SetValidationList(_cells.GetCell(5, titleRow + 1), ValidationSheetName, 10, false);
                _cells.GetCell(6, titleRow).Value = "OUT_SERV_DATE";
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow).Value = "COMENTARIO";
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow).AddComment("Adiciona el siguiente texto al campo de comentario (no elimina el comentario existente)");

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 5 - CLOSE COMMENTS
                titleRow = TitleRow05;
                resultColumn = ResultColumn05;
                tableName = TableName05;
                sheetName = SheetName05;
                _excelApp.ActiveWorkbook.Sheets[5].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

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

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, titleRow).Value = "WORK_ORDER";
                _cells.GetCell(2, titleRow).Value = "COMENTARIO";
                _cells.GetCell(3, titleRow).Value = "COMPLETED_DATE";
                _cells.GetCell(4, titleRow).Value = "COMPLETED_BY";
                _cells.GetCell(2, titleRow).Style = StyleConstants.TitleOptional;

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 6 - DURATION
                titleRow = TitleRow06;
                resultColumn = ResultColumn06;
                tableName = TableName06;
                sheetName = SheetName06;
                _excelApp.ActiveWorkbook.Sheets[6].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "WORK ORDERS DURATIONS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                //GENERAL
                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(1, titleRow).Value = "DISTRICT_CODE";
                _cells.GetCell(2, titleRow).Value = "WORK_ORDER";
                _cells.GetCell(3, titleRow).Value = "DURATION_DATE";
                _cells.GetCell(3, titleRow).AddComment("yyyyMMdd");
                _cells.GetCell(4, titleRow).Value = "DURATION_CODE";
                var durationCodeList = _eFunctions.GetItemCodesString("JI");
                _cells.SetValidationList(_cells.GetCell(4, titleRow + 1), durationCodeList, ValidationSheetName, 11, false);
                _cells.GetCell(5, titleRow).Value = "START_HOUR";
                _cells.GetCell(5, titleRow).AddComment("hhmmss");
                _cells.GetCell(6, titleRow).Value = "FINAL_HOUR";
                _cells.GetCell(6, titleRow).AddComment("hhmmss");
                _cells.GetCell(7, titleRow).Value = "DURATION_TIME";
                _cells.GetCell(7, titleRow).AddComment("En formato numérico. Ej. 2.5 horas (000000 - 023000) hhmmss");
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(8, titleRow).Value = "ACTION";
                _cells.GetCell(8, titleRow).Style = StyleConstants.TitleAction;
                _cells.GetCell(8, titleRow).AddComment("Crear, Eliminar");
                var actionsList = new List<string> { "Crear", "Eliminar" };
                _cells.SetValidationList(_cells.GetCell(8, titleRow + 1), actionsList, ValidationSheetName, 12, false);

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 7 - PROGRESS WO
                titleRow = TitleRow07;
                resultColumn = ResultColumn07;
                tableName = TableName07;
                sheetName = SheetName07;
                _excelApp.ActiveWorkbook.Sheets[7].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

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
                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(1, titleRow).Value = "WORK_ORDER";
                _cells.GetCell(2, titleRow).Value = "UNITS OF WORK";
                _cells.GetCell(3, titleRow).Value = "UNITS REQUIRED";
                _cells.GetCell(4, titleRow).Value = "PERCENT COMPLETED";
                _cells.GetCell(5, titleRow).Value = "UNITS COMPLETED";

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion
                
                #region CONSTRUYO LA HOJA 8 - TO DO LIST
                titleRow = TitleRow08;
                resultColumn = ResultColumn08;
                _excelApp.ActiveWorkbook.Sheets[8].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName08;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "WORK ORDERS TO DO - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                //GENERAL
                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(1, titleRow).Value = "DISTRICT";
                _cells.GetCell(2, titleRow).Value = "WORK ORDER";
                _cells.GetCell(3, titleRow).Value = "WO. TASK";
                _cells.GetCell(4, titleRow).Value = "SEQUENCE";
                _cells.GetCell(5, titleRow).Value = "ITEM NAME";
                _cells.GetCell(6, titleRow).Value = "REQ. DATE";
                _cells.GetCell(7, titleRow).Value = "EXP. DATE";
                _cells.GetCell(8, titleRow).Value = "NEED FOR RELEASE";
                _cells.GetCell(9, titleRow).Value = "EXT. REFERENCE";
                _cells.GetCell(10, titleRow).Value = "OWNER";
                _cells.GetCell(11, titleRow).Value = "NOTES";
                _cells.GetCell(12, titleRow).Value = "STATUS";

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
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

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
                var usTypeCodes = new List<string>();
                foreach (var item in WorkOrderActions.GetUserStatusCodeList(_eFunctions))
                    usTypeCodes.Add(item.Code + " - " + item.Description);
                var contactMethod = _eFunctions.GetItemCodesString("MTCO");


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
                _cells.GetCell(35, TitleRowD04).Value = "Calificacion Calidad";
                _cells.GetCell(36, TitleRowD04).Value = "Calificado Por";
                _cells.GetCell(37, TitleRowD04).Value = "Secuencia OT";

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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opContext = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opContext = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                }
            }
            _eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void CreateWoDetailedList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameD01, ResultColumnD01);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowD01 + 1;
            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDstrct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);

            while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)) || !string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value2)))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    //GENERAL
                    wo.districtCode = _frmAuth.EllipseDstrct;
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
                    wo.calculatedDurationsFlag = cbFlagEstDuration.Checked.ToString();

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
                    var replySheet = WorkOrderActions.CreateWorkOrder(urlService, opSheet, wo);
                    wo.SetWorkOrderDto(replySheet.workOrder.prefix, replySheet.workOrder.no);
                    if (cbFlagEstDuration.Checked)
                    {
                        _cells.GetCell(28, i).Value = replySheet.planStrDate;
                        _cells.GetCell(29, i).Value = replySheet.planStrTime;
                        _cells.GetCell(30, i).Value = replySheet.planFinDate;
                        _cells.GetCell(31, i).Value = replySheet.planFinTime;
                    }
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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRowD01 + 1;
            const int validationRow = TitleRowD01 - 1;

            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDstrct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);


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
                    wo.calculatedDurationsFlag = cbFlagEstDuration.Checked.ToString();

                    var reply = WorkOrderActions.ModifyWorkOrder(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, wo);
                    if (cbFlagEstDuration.Checked)
                    {
                        _cells.GetCell(28, i).Value = reply.planStrDate;
                        _cells.GetCell(29, i).Value = reply.planStrTime;
                        _cells.GetCell(30, i).Value = reply.planFinDate;
                        _cells.GetCell(31, i).Value = reply.planFinTime;
                    }

                    var extendedHeader = MyUtilities.IsTrue(_cells.GetCell(54, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(54, i).Value) : null;
                    var extendedBody = MyUtilities.IsTrue(_cells.GetCell(55, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(55, i).Value) : null;
                    wo.SetExtendedDescription(extendedHeader, extendedBody);

                    WorkOrderActions.ModifyWorkOrder(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, wo);
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
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

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
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
        }
        private void ReviewWoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName01);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

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

            if (string.IsNullOrWhiteSpace(searchCriteriaValue1) && string.IsNullOrWhiteSpace(searchCriteriaValue2))
            {
                MessageBox.Show("Debe seleccionar al menos una opción de búsqueda", "Error de Consulta");
                if (_cells != null) _cells.SetCursorDefault();
                return;
            }
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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
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
                }
            }
			_eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReviewQualityList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableNameQ01);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
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

            var completeCodeList = _eFunctions.GetItemCodesDictionary("SC");

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
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
                var completeCodeList = _eFunctions.GetItemCodesDictionary("SC");
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
                        _cells.GetCell(23, i).Value =
                            "" + (wo.calculatedLabFlag.Equals("Y") ? wo.calculatedLabHrs : wo.estimatedLabHrs);
                        _cells.GetCell(24, i).Value2 = "" + wo.actualLabHrs;
                        _cells.GetCell(24, i).Style = wc.ActualLabHrs;
                        _cells.GetCell(25, i).Value2 = "" + (wo.calculatedLabFlag.Equals("Y")
                            ? wo.calculatedLabCost
                            : wo.estimatedLabCost);
                        _cells.GetCell(26, i).Value2 = "" + wo.actualLabCost;
                        _cells.GetCell(26, i).Style = wc.ActualLabCost;
                        _cells.GetCell(27, i).Value2 = "" + (wo.calculatedMatFlag.Equals("Y")
                            ? wo.calculatedMatCost
                            : wo.estimatedMatCost);
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

                        _cells.GetCell(33, i).Value = completeCodeList.ContainsKey(wo.completedCode)
                            ? "'" + wo.completedCode + " - " +
                              completeCodeList.TryGetValue(wo.completedCode, out outString)
                            : "";
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
                    }
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReReviewWOList()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Error. " + ex.Message);
            }
            finally
            {
                _eFunctions.CloseConnection();
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
            }

        }
        private void CreateWoList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDstrct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);

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
                    //se está forzando porque recientemente en una actualización de E8, si no se envía (se envía nulo) el predeterminado es falso
                    wo.calculatedEquipmentFlag = "true";
                    wo.calculatedMatFlag = "true";
                    wo.calculatedOtherFlag = "true";
                    wo.calculatedLabFlag = "true";
                    wo.calculatedDurationsFlag = cbFlagEstDuration.Checked.ToString();

                    var replySheet = WorkOrderActions.CreateWorkOrder(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, wo);
                    if (cbFlagEstDuration.Checked)
                    {
                        _cells.GetCell(28, i).Value = replySheet.planStrDate;
                        _cells.GetCell(29, i).Value = replySheet.planStrTime;
                        _cells.GetCell(30, i).Value = replySheet.planFinDate;
                        _cells.GetCell(31, i).Value = replySheet.planFinTime;
                    }
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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDstrct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();
            var planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWO3", EllipseFunctions.ProgramAccessType.Full);
            if (!planValidation)
                planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWJO", EllipseFunctions.ProgramAccessType.Full);
            //if (!planValidation)
            //    planValidation = _eFunctions.CheckUserProgramAccess(drpEnvironment.SelectedItem.Label, district, userName, "MSEWOT", EllipseFunctions.ProgramAccessType.Full);


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
                    wo.calculatedDurationsFlag = cbFlagEstDuration.Checked.ToString();

                    var reply = WorkOrderActions.ModifyWorkOrder(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, wo);
                    if(cbFlagEstDuration.Checked)
                    {
                        _cells.GetCell(28, i).Value = reply.planStrDate;
                        _cells.GetCell(29, i).Value = reply.planStrTime;
                        _cells.GetCell(30, i).Value = reply.planFinDate;
                        _cells.GetCell(31, i).Value = reply.planFinTime;
                    }

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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName04, ResultColumn04);

            var i = TitleRow04 + 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                        if(woData == null)
                            throw new Exception("La orden " + wo.workOrder.prefix + wo.workOrder.no + " no existe o no ha sido encontrada");
                        if (WoStatusList.ClosedCode.Equals(woData.workOrderStatusM.Trim()) || WoStatusList.CancelledCode.Equals(woData.workOrderStatusM.Trim()))
                            throw new Exception("La orden " + wo.workOrder.prefix + wo.workOrder.no + " ya está cerrada como " + WoStatusList.GetStatusName(woData.workOrderStatusM.Trim()) + " con código " + woData.completedCode);
                    }
                    var reply = WorkOrderActions.CompleteWorkOrder(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, wo);
                    if (reply.completedCode.Trim().Equals(wo.completedCode.Trim(), StringComparison.InvariantCultureIgnoreCase) && reply.closedDate == wo.closedDate)
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
            _eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
            
        }
        private void ReOpenWoList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName04, ResultColumn04);

            var i = TitleRow04 + 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                        if (!WoStatusList.ClosedCode.Equals(woData.workOrderStatusM.Trim()) && !WoStatusList.CancelledCode.Equals(woData.workOrderStatusM.Trim()))
                            throw new Exception("La orden " + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no + " ya está abierta como " + WoStatusList.GetStatusName(woData.workOrderStatusM.Trim()));
                    }
                    WorkOrderActions.ReOpenWorkOrder(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, wo);

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
			_eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void ReviewCloseText()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
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
                    var closeText = WorkOrderActions.GetWorkOrderCloseText(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), districtCode, _frmAuth.EllipsePost, Debugger.DebugWarnings, wo);
                    var workOrder = WorkOrderActions.FetchWorkOrder(_eFunctions, districtCode, wo);

                    _cells.GetCell(2, i).Value = closeText;
                    _cells.GetCell(3, i).Value = workOrder.closeCommitDate; 
                    _cells.GetCell(4, i).Value = workOrder.completedBy;
                    _cells.GetCell(ResultColumn05, i).Value = "OK - " + workOrder.completedCode;
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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName05, ResultColumn05);

            var i = TitleRow05 + 1;

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                    //WorkOrderActions.SetWorkOrderCloseText(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), districtCode, _frmAuth.EllipsePost, Debugger.DebugWarnings, wo, closeText);
                    WorkOrderActions.AppendTextToCloseComment(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), opSheet, districtCode, wo.prefix + wo.no, closeText);

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
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var opSheet = new WorkOrderService.OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var i = TitleRow01 + 1;
                var k = TitleRow06 + 1;
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName06, ResultColumn06);

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow06 + 1;
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            while (!string.IsNullOrWhiteSpace(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value)))
            {
                try
                {
                    string districtCode = _cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value);
                    if (string.IsNullOrWhiteSpace(districtCode))
                        districtCode = "ICOR";

                    var wo = WorkOrderActions.GetNewWorkOrderDto(_cells.GetEmptyIfNull(_cells.GetCell(2, i).Value));

                    //Corrección de Start Hour y Finish Hour
                    string startTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    string finishTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                    string stringHoursTime = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                    
                    //Si solo se ingresó la hora final se asume la inicial como la hora 0
                    if (string.IsNullOrWhiteSpace(startTime) && !string.IsNullOrWhiteSpace(finishTime))
                        startTime = "000000";
                    //si se ingresa la información con solo la duración sin especificar las horas
                    else if(string.IsNullOrWhiteSpace(startTime) && string.IsNullOrWhiteSpace(finishTime) && !string.IsNullOrWhiteSpace(stringHoursTime))
                    {
                        startTime = "000000";
                        finishTime = MyUtilities.DateTime.ConvertDecimalHourToHHMM(stringHoursTime, "") + "00";
                    }

                    //Ellipse presenta problemas cuando los segundos ingresados son diferentes de 0
                    //Se restringe para que siempre los asuma como 0
                    if (startTime != null && startTime.Length == 6)
                        startTime = startTime.Substring(0, 4) + "00";
                    if (finishTime != null && finishTime.Length == 6)
                        finishTime = finishTime.Substring(0, 4) + "00";
                    //

                    var duration = new WorkOrderDuration
                    {
                        jobDurationsDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value),
                        jobDurationsCode = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value)),
                        jobDurationsStart = startTime,
                        jobDurationsFinish = finishTime
                    };
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value).ToUpper();
                    switch (action)
                    {
                        case "CREAR":
                            {
                                WorkOrderActions.CreateWorkOrderDuration(urlService, opSheet, districtCode, wo, duration);
                                _cells.GetCell(ResultColumn06, i).Value = "CREADO";
                                _cells.GetCell(ResultColumn06, i).Style = StyleConstants.Success;
                                _cells.GetCell(ResultColumn06 - 1, i).Value = "";//Para evitar duplicados por repetición
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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                    _cells.GetCell(35, i).Value = "'" + woRefCodes.CalificacionCalidadOt;
                    _cells.GetCell(36, i).Value = "'" + woRefCodes.CalificacionCalidadPor;
                    _cells.GetCell(37, i).Value = "'" + woRefCodes.SecuenciaOt;

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
                }
            }
			_eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }
        private void UpdateReferenceCodes()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameD04, ResultColumnD04);

            var i = TitleRowD04 + 1;
            const int validationRow = TitleRowD04 - 1;

            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                        MetodoContactoText = MyUtilities.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null,
                        CalificacionCalidadOt =  MyUtilities.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null,
                        CalificacionCalidadPor =  MyUtilities.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null,
                        SecuenciaOt =  MyUtilities.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null
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
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
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
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var stOpContext = StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, true);
            _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
            var woCells = new ExcelStyleCells(_excelApp, SheetName01);
            woCells.SetFixedWorkingWorkSheet(true);

            var j = TitleRow01 + 1;//itera según cada estándar
            var i = TitleRow02 + 1;//itera la celda para cada tarea

            while (!string.IsNullOrEmpty("" + woCells.GetCell(3, j).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(woCells.GetCell(2, 3).Value2);
                    var workOrder = _cells.GetEmptyIfNull(woCells.GetCell(2, j).Value2);
                    var woPlanStartDate = _cells.GetEmptyIfNull(woCells.GetCell(28, j).Value2);
                    var woPlanStartTime = _cells.GetEmptyIfNull(woCells.GetCell(29, j).Value2);
                    var woPlanFinishDate = _cells.GetEmptyIfNull(woCells.GetCell(30, j).Value2);
                    var woPlanFinishTime = _cells.GetEmptyIfNull(woCells.GetCell(31, j).Value2);

                    woPlanStartDate = string.IsNullOrWhiteSpace(woPlanStartDate) ? "000000" : woPlanStartDate.PadLeft(8, '0'); ;
                    woPlanStartTime = string.IsNullOrWhiteSpace(woPlanStartTime) ? "000000" : woPlanStartTime.PadLeft(6, '0'); ;
                    woPlanFinishDate = string.IsNullOrWhiteSpace(woPlanFinishDate) ? "000000" : woPlanFinishDate.PadLeft(8, '0'); ;
                    woPlanFinishTime = string.IsNullOrWhiteSpace(woPlanFinishTime) ? "000000" : woPlanFinishTime.PadLeft(6, '0'); ;

                    var taskList = WorkOrderTaskActions.FetchWorkOrderTask(_eFunctions, districtCode, workOrder, "");


                    foreach (WorkOrderTask task in taskList)
                    {
                        //Para resetear el estilo
                        _cells.GetRange(1, i, ResultColumn02, i).Style = StyleConstants.Normal;
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + task.DistrictCode;
                        _cells.GetCell(2, i).Value = "" + task.WorkGroup;
                        _cells.GetCell(3, i).Value = "'" + task.WorkOrder;
                        _cells.GetCell(4, i).Value = "" + task.WorkOrderDescription;
                        //ACTION
                        _cells.GetCell(5, i).Value = WorkOrderTaskActions.Modify;
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
                        _cells.GetCell(15, i).Value = "'" + task.PlanStartTime; 
                        _cells.GetCell(16, i).Value = "'" + task.PlanFinishDate;
                        _cells.GetCell(17, i).Value = "'" + task.PlanFinishTime;
                        //Valida que las fechas de plan de la tarea estén dentro de las fechas de la orden
                        if (cbValidateTaskPlanDates.Checked && (!string.IsNullOrWhiteSpace(woPlanStartDate) && !string.IsNullOrWhiteSpace(woPlanFinishDate)))
                        {
                            _cells.GetRange(14, i, 17, i).ClearComments();
                            var tkPlanStartDate = string.IsNullOrWhiteSpace(task.PlanStartDate) ? "000000" : task.PlanStartDate.PadLeft(8, '0');
                            var tkPlanStartTime = string.IsNullOrWhiteSpace(task.PlanStartTime) ? "000000" : task.PlanStartTime.PadLeft(6, '0');
                            var tkPlanFinishDate = string.IsNullOrWhiteSpace(task.PlanFinishDate) ? "000000" : task.PlanFinishDate.PadLeft(8, '0');
                            var tkPlanFinishTime = string.IsNullOrWhiteSpace(task.PlanFinishTime) ? "000000" : task.PlanFinishTime.PadLeft(6, '0');

                            if (!string.IsNullOrWhiteSpace(task.PlanStartDate))
                            {
                                if ((Convert.ToDouble(string.Concat(tkPlanStartDate,tkPlanStartTime)) < Convert.ToDouble(string.Concat(woPlanStartDate, woPlanStartTime))) || (Convert.ToDouble(string.Concat(tkPlanStartDate, tkPlanStartTime)) > Convert.ToDouble(string.Concat(woPlanFinishDate, woPlanFinishTime))))
                                {
                                    _cells.GetCell(14, i).Style = StyleConstants.Error;
                                    _cells.GetCell(14, i).AddComment("WoPlanStartDate: " + woPlanStartDate + " " + woPlanStartTime + "\nWoPlanFinishDate: " + woPlanFinishDate + " " + woPlanFinishTime);
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(task.PlanFinishDate))
                            {
                                if ((Convert.ToDouble(string.Concat(tkPlanFinishDate, tkPlanFinishTime)) < Convert.ToDouble(string.Concat(woPlanStartDate, woPlanStartTime))) || (Convert.ToDouble(string.Concat(tkPlanFinishDate, tkPlanFinishTime)) > Convert.ToDouble(string.Concat(woPlanFinishDate, woPlanFinishTime))))
                                {
                                    _cells.GetCell(16, i).Style = StyleConstants.Error;
                                    _cells.GetCell(16, i).AddComment("WoPlanStartDate: " + woPlanStartDate + " " + woPlanStartTime + "\nWoPlanFinishDate: " + woPlanFinishDate + " " + woPlanFinishTime);
                                }
                            }
                        }
                        //RECURSOS
                        _cells.GetCell(18, i).Value = "" + task.EstimatedDurationsHrs;
                        _cells.GetCell(19, i).Value = "" + task.NoLabor;
                        _cells.GetCell(20, i).Value = "" + task.NoMaterial;
                        //APL
                        _cells.GetCell(21, i).Value = "'" + task.AplEquipmentGrpId;
                        _cells.GetCell(22, i).Value = "'" + task.AplType;
                        _cells.GetCell(23, i).Value = "'" + task.AplCompCode;
                        _cells.GetCell(24, i).Value = "'" + task.AplCompModCode;
                        _cells.GetCell(25, i).Value = "'" + task.AplSeqNo;
                        _cells.GetRange(21, i, 25, i).Style = !string.IsNullOrWhiteSpace(task.AplType) ? StyleConstants.Error : StyleConstants.Normal;

                        var stdTextId = "WI" + task.DistrictCode + task.WorkOrder + task.WoTaskNo;
                        _cells.GetCell(26, i).Value = StdText.GetText(urlService, stOpContext, stdTextId);
                        _cells.GetCell(26, i).WrapText = false;
                        _cells.GetCell(27, i).Value = "'" + task.CompletedCode;
                        _cells.GetCell(28, i).Value = "'" + task.CompletedBy;
                        _cells.GetCell(29, i).Value = "'" + task.ClosedDate;
                        var stdTextCompleteId = "WA" + task.DistrictCode + task.WorkOrder + task.WoTaskNo;
                        _cells.GetCell(30, i).Value = StdText.GetText(urlService, stOpContext, stdTextCompleteId);
                        _cells.GetCell(30, i).WrapText = false;
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
                }
            }
			_eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        private void CreateWoToDoList()
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            var tableName = TableName08;
            var resultColumn = ResultColumn08;
            var titleRow = TitleRow08;

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var opContext = WorkOrderToDoActions.GetOperationContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost);

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty(_cells.GetCell(2, i).Value2))
            {
                try
                {
                    var toDo = new WorkOrderToDoItem();
                          
                    toDo.DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    toDo.WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    toDo.WorkOrderTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    toDo.Sequence = !string.IsNullOrWhiteSpace(_cells.GetCell(4, i).Value) ? Convert.ToDecimal(_cells.GetCell(4, i).Value) : default(decimal);
                    toDo.SequenceSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(4, i).Value);
                    toDo.ItemName = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    toDo.RequiredByDate = MyUtilities.ToDateTime(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value));
                    toDo.RequiredByDateSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(6, i).Value);
                    toDo.ExpirationDate = MyUtilities.ToDateTime(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    toDo.ExpirationDateSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(7, i).Value);
                    toDo.NeededForRelease = MyUtilities.IsTrue(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    toDo.NeededForReleaseSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(8, i).Value);
                    toDo.ExternalReference = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value);
                    toDo.Owner = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                    toDo.Notes = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    toDo.StatusCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);

                    var replyItem = WorkOrderToDoActions.CreateToDoItems(urlService, opContext, toDo);

                    _cells.GetCell(1, i).Value = "" + replyItem.DistrictCode;
                    _cells.GetCell(2, i).Value = "'" + replyItem.WorkOrder;
                    _cells.GetCell(3, i).Value = "'" + replyItem.WorkOrderTask;
                    _cells.GetCell(4, i).Value = "'" + replyItem.Sequence;
                    _cells.GetCell(5, i).Value = "'" + replyItem.ItemName;
                    _cells.GetCell(6, i).Value = "'" + MyUtilities.ToString(replyItem.RequiredByDate);
                    _cells.GetCell(7, i).Value = "'" + MyUtilities.ToString(replyItem.ExpirationDate);
                    _cells.GetCell(8, i).Value = "'" + (replyItem.NeededForRelease ? "Y" : "N");
                    _cells.GetCell(9, i).Value = "'" + replyItem.ExternalReference;
                    _cells.GetCell(10, i).Value = "'" + replyItem.Owner;
                    _cells.GetCell(11, i).Value = "'" + replyItem.Notes;
                    _cells.GetCell(12, i).Value = "'" + replyItem.StatusCode;

                    _cells.GetCell(resultColumn, i).Value = "TO DO CREADO";
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateWoToDoList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void DeleteWoToDoList()
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            var tableName = TableName08;
            var resultColumn = ResultColumn08;
            var titleRow = TitleRow08;

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var opContext = WorkOrderToDoActions.GetOperationContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost);

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value2))
            {
                try
                {
                    var toDo = new WorkOrderToDoItem();

                    toDo.DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    toDo.WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    toDo.WorkOrderTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    toDo.Sequence = !string.IsNullOrWhiteSpace(_cells.GetCell(4, i).Value) ? Convert.ToDecimal(_cells.GetCell(4, i).Value) : default(decimal);
                    toDo.SequenceSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(4, i).Value);
                    toDo.ItemName = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    toDo.RequiredByDate = MyUtilities.ToDateTime(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value));
                    toDo.RequiredByDateSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(6, i).Value);
                    toDo.ExpirationDate = MyUtilities.ToDateTime(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    toDo.ExpirationDateSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(7, i).Value);
                    toDo.NeededForRelease = MyUtilities.IsTrue(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    toDo.NeededForReleaseSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(8, i).Value);
                    toDo.ExternalReference = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value);
                    toDo.Owner = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                    toDo.Notes = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                    toDo.StatusCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);

                    var replyItem = WorkOrderToDoActions.DeleteToDoItems(urlService, opContext, toDo);

                    _cells.GetCell(1, i).Value = "" + replyItem.DistrictCode;
                    _cells.GetCell(2, i).Value = "'" + replyItem.WorkOrder;
                    _cells.GetCell(3, i).Value = "'" + replyItem.WorkOrderTask;
                    _cells.GetCell(4, i).Value = "'" + replyItem.Sequence;
                    _cells.GetCell(5, i).Value = "'" + replyItem.ItemName;
                    _cells.GetCell(6, i).Value = "'" + MyUtilities.ToString(replyItem.RequiredByDate);
                    _cells.GetCell(7, i).Value = "'" + MyUtilities.ToString(replyItem.ExpirationDate);
                    _cells.GetCell(8, i).Value = "'" + (replyItem.NeededForRelease ? "Y" : "N");
                    _cells.GetCell(9, i).Value = "'" + replyItem.ExternalReference;
                    _cells.GetCell(10, i).Value = "'" + replyItem.Owner;
                    _cells.GetCell(11, i).Value = "'" + replyItem.Notes;
                    _cells.GetCell(12, i).Value = "'" + replyItem.StatusCode;

                    _cells.GetCell(resultColumn, i).Value = "ELIMINADO";
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:DeleteWoToDoList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ReviewWoToDo()
        {
            var titleRow = TitleRow08;
            var resultColumn = ResultColumn08;
            var tableName = TableName08;

            //var resultColumnJ = ResultColumn01;
            //var tableNameJ = TableName01;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(tableName);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opContext = WorkOrderToDoActions.GetOperationContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var woCells = new ExcelStyleCells(_excelApp, SheetName01);
            var tdCells = new ExcelStyleCells(_excelApp, SheetName08);
            woCells.SetFixedWorkingWorkSheet(true);
            tdCells.SetFixedWorkingWorkSheet(true);

            var j = TitleRow01 + 1;//itera según cada orden
            var i = titleRow + 1;//itera la celda para to do
            //reubicación del cursor en la última fila disponible
            while (!string.IsNullOrWhiteSpace("" + tdCells.GetCell(2, i).Value))
                i++;

            _excelApp.ActiveWorkbook.Sheets[8].Select(Type.Missing);

            while (!string.IsNullOrEmpty("" + woCells.GetCell(2, j).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(woCells.GetCell(2, 3).Value2);
                    var workOrder = _cells.GetEmptyIfNull(woCells.GetCell(2, j).Value2);
                    string workOrderTask = null;
                    var toDoList = WorkOrderToDoActions.FetchToDoItems(urlService, districtCode, workOrder, workOrderTask, opContext);


                    foreach (var toDo in toDoList)
                    {
                        //Para resetear el estilo
                        tdCells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        tdCells.GetCell(1, i).Value = "" + toDo.DistrictCode;
                        tdCells.GetCell(2, i).Value = "'" + toDo.WorkOrder;
                        tdCells.GetCell(3, i).Value = "'" + toDo.WorkOrderTask;
                        tdCells.GetCell(4, i).Value = "'" + toDo.Sequence;
                        tdCells.GetCell(5, i).Value = "'" + toDo.ItemName;
                        tdCells.GetCell(6, i).Value = "'" + MyUtilities.ToString(toDo.RequiredByDate);
                        tdCells.GetCell(7, i).Value = "'" + MyUtilities.ToString(toDo.ExpirationDate);
                        tdCells.GetCell(8, i).Value = "'" + (toDo.NeededForRelease ? "Y" : "N");
                        tdCells.GetCell(9, i).Value = "'" + toDo.ExternalReference;
                        tdCells.GetCell(10, i).Value = "'" + toDo.Owner;
                        tdCells.GetCell(11, i).Value = "'" + toDo.Notes;
                        tdCells.GetCell(12, i).Value = "'" + toDo.StatusCode;

                        tdCells.GetCell(4, i).Select();
                        i++;//aumenta to do
                    }
                }
                catch (Exception ex)
                {
                    tdCells.GetCell(1, i).Style = StyleConstants.Error;
                    tdCells.GetCell(1, i).Value = _cells.GetEmptyIfNull(woCells.GetCell(2, 3).Value2);
                    tdCells.GetCell(2, i).Value = _cells.GetEmptyIfNull(woCells.GetCell(2, j).Value2);
                    tdCells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    tdCells.GetCell(resultColumn, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:ReviewWoToDo()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta wo
                }
            }
			_eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        private void ReviewWoTaskToDo()
        {
            var titleRow = TitleRow08;
            var resultColumn = ResultColumn08;
            var tableName = TableName08;

            //var resultColumnJ = ResultColumn01;
            //var tableNameJ = TableName01;

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(tableName);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var opContext = WorkOrderToDoActions.GetOperationContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var tkCells = new ExcelStyleCells(_excelApp, SheetName02);
            var tdCells = new ExcelStyleCells(_excelApp, SheetName08);
            tkCells.SetFixedWorkingWorkSheet(true);
            tdCells.SetFixedWorkingWorkSheet(true);

            var j = TitleRow02 + 1;//itera según cada orden
            var i = titleRow + 1;//itera la celda para to do
            //reubicación del cursor en la última fila disponible
            while (!string.IsNullOrWhiteSpace("" + tdCells.GetCell(2, i).Value))
                i++;

            _excelApp.ActiveWorkbook.Sheets[8].Select(Type.Missing);

            while (!string.IsNullOrEmpty("" + tkCells.GetCell(3, j).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(tkCells.GetCell(1, j).Value2);
                    var workOrder = _cells.GetEmptyIfNull(tkCells.GetCell(3, j).Value2);
                    var workOrderTask = _cells.GetEmptyIfNull(tkCells.GetCell(6, j).Value2);

                    var toDoList = WorkOrderToDoActions.FetchToDoItems(urlService, districtCode, workOrder, workOrderTask, opContext);


                    foreach (var toDo in toDoList)
                    {
                        //Para resetear el estilo
                        tdCells.GetRange(1, i, resultColumn, i).Style = StyleConstants.Normal;

                        tdCells.GetCell(1, i).Value = "" + toDo.DistrictCode;
                        tdCells.GetCell(2, i).Value = "'" + toDo.WorkOrder;
                        tdCells.GetCell(3, i).Value = "'" + toDo.WorkOrderTask;
                        tdCells.GetCell(4, i).Value = "'" + toDo.Sequence;
                        tdCells.GetCell(5, i).Value = "'" + toDo.ItemName;
                        tdCells.GetCell(6, i).Value = "'" + MyUtilities.ToString(toDo.RequiredByDate);
                        tdCells.GetCell(7, i).Value = "'" + MyUtilities.ToString(toDo.ExpirationDate);
                        tdCells.GetCell(8, i).Value = "'" + (toDo.NeededForRelease ? "Y" : "N");
                        tdCells.GetCell(9, i).Value = "'" + toDo.ExternalReference;
                        tdCells.GetCell(10, i).Value = "'" + toDo.Owner;
                        tdCells.GetCell(11, i).Value = "'" + toDo.Notes;
                        tdCells.GetCell(12, i).Value = "'" + toDo.StatusCode;

                        tdCells.GetCell(4, i).Select();
                        i++;//aumenta to do
                    }
                }
                catch (Exception ex)
                {
                    tdCells.GetCell(1, i).Style = StyleConstants.Error;
                    tdCells.GetCell(1, i).Value = _cells.GetEmptyIfNull(tkCells.GetCell(2, 3).Value2);
                    tdCells.GetCell(2, i).Value = _cells.GetEmptyIfNull(tkCells.GetCell(2, j).Value2);
                    tdCells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    tdCells.GetCell(resultColumn, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:ReviewWoTaskToDo()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta wo
                }
            }
            _eFunctions.CloseConnection();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        private void UpdateWoToDoList()
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            var tableName = TableName08;
            var resultColumn = ResultColumn08;
            var titleRow = TitleRow08;

            _cells.ClearTableRangeColumn(tableName, resultColumn);

            var i = titleRow + 1;
            var opContext = WorkOrderToDoActions.GetOperationContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost);

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrWhiteSpace("" + _cells.GetCell(2, i).Value2))
            {
                try
                {
                    var toDo = new WorkOrderToDoItem();

                    toDo.DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    toDo.WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    toDo.WorkOrderTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    toDo.Sequence = !string.IsNullOrWhiteSpace(_cells.GetCell(4, i).Value) ? Convert.ToDecimal(_cells.GetCell(4, i).Value) : default(decimal);
                    toDo.SequenceSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(4, i).Value);
                    toDo.ItemName = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    toDo.RequiredByDate = MyUtilities.ToDateTime(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value));
                    toDo.RequiredByDateSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(6, i).Value);
                    toDo.ExpirationDate = MyUtilities.ToDateTime(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    toDo.ExpirationDateSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(7, i).Value);
                    toDo.NeededForRelease = MyUtilities.IsTrue(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    toDo.NeededForReleaseSpecified = !string.IsNullOrWhiteSpace(_cells.GetCell(8, i).Value);
                    toDo.ExternalReference = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value);
                    toDo.Owner = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value);
                    toDo.Notes = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value);
                    toDo.StatusCode = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value);

                    var replyItem = WorkOrderToDoActions.UpdateToDoItems(urlService, opContext, toDo);

                    _cells.GetCell(1, i).Value = "" + replyItem.DistrictCode;
                    _cells.GetCell(2, i).Value = "'" + replyItem.WorkOrder;
                    _cells.GetCell(3, i).Value = "'" + replyItem.WorkOrderTask;
                    _cells.GetCell(4, i).Value = "'" + replyItem.Sequence;
                    _cells.GetCell(5, i).Value = "'" + replyItem.ItemName;
                    _cells.GetCell(6, i).Value = "'" + MyUtilities.ToString(replyItem.RequiredByDate);
                    _cells.GetCell(7, i).Value = "'" + MyUtilities.ToString(replyItem.ExpirationDate);
                    _cells.GetCell(8, i).Value = "'" + (replyItem.NeededForRelease ? "Y" : "N");
                    _cells.GetCell(9, i).Value = "'" + replyItem.ExternalReference;
                    _cells.GetCell(10, i).Value = "'" + replyItem.Owner;
                    _cells.GetCell(11, i).Value = "'" + replyItem.Notes;
                    _cells.GetCell(12, i).Value = "'" + replyItem.StatusCode;

                    _cells.GetCell(resultColumn, i).Value = "ACTUALIZADO";
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                    _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateWoToDoList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(resultColumn, i).Select();
                    i++;
                }
            }
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void btnReviewRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewRequirements(RequirementType.All.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }


        private void ReviewRequirements(string requirementType)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
            _cells.ClearTableRange(TableName03);
            var woCells = new ExcelStyleCells(_excelApp, SheetName01);
            woCells.SetFixedWorkingWorkSheet(true);

            var resultColumn = ResultColumn03;
            var j = TitleRow01 + 1;//itera según cada orden
            var i = TitleRow03 + 1;//itera la celda para cada requerimiento

            var list = new List<TaskRequirement>();

            while (!string.IsNullOrEmpty("" + woCells.GetCell(3, j).Value))
            {
                list.Add(new TaskRequirement
                {
                    DistrictCode = _cells.GetEmptyIfNull(woCells.GetCell("B3").Value2),
                    WorkGroup = _cells.GetEmptyIfNull(woCells.GetCell(1, j).Value2),
                    WorkOrder = _cells.GetEmptyIfNull(woCells.GetCell(2, j).Value2),
                });
                j++;
            }

            var distinctItems = list.GroupBy(x => new { x.DistrictCode, x.WorkGroup, x.WorkOrder}).Select(y => y.First());

            foreach (var d in distinctItems)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(d.DistrictCode))
                        d.DistrictCode = "ICOR";
                    var reqList = WorkOrderTaskActions.FetchRequirements(_eFunctions, d.DistrictCode, d.WorkOrder, requirementType, null);

                    var distinctReqList = reqList.GroupBy(x => new { x.DistrictCode, x.WorkGroup, x.WorkOrder, x.WoTaskNo, x.ReqCode }).Select(y => y.First());

                    foreach (var req in distinctReqList)
                    {
                        _cells.GetRange(1, i, resultColumn, i).ClearFormats();
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + req.DistrictCode; //DistrictCode
                        _cells.GetCell(2, i).Value = "" + req.WorkGroup;    //WorkGroup
                        _cells.GetCell(3, i).Value = "'" + req.WorkOrder;    //WorkOrder 
                        _cells.GetCell(4, i).Value = "'" + req.WoTaskNo;     //WoTaskNo 
                        _cells.GetCell(5, i).Value = "" + req.WoTaskDesc;   //WoTaskDesc 
                        _cells.GetCell(6, i).Value = "M";
                        _cells.GetCell(7, i).Value = "" + req.ReqType;      //ReqType 
                        _cells.GetCell(8, i).Value = "" + req.SeqNo;        //SeqNo 
                        _cells.GetCell(9, i).Value = "" + req.ReqCode;      //ReqCode
                        _cells.GetCell(10, i).Value = "" + req.ReqDesc;     //ReqDesc
                        _cells.GetCell(11, i).Value = "" + req.UoM;         //UoM
                        _cells.GetCell(12, i).Value = "" + req.EstSize;      //EstSize
                        _cells.GetCell(13, i).Value = "" + req.UnitsQty;      //UnitsQty
                        _cells.GetCell(14, i).Value = "" + req.RealQty;     //RealQty
                        _cells.GetCell(15, i).Value = "" + req.SharedTasks;     //SharedTask
                        if (Convert.ToInt16(req.SharedTasks) != 1)
                            _cells.GetRange(14, i, 15, i).Style = StyleConstants.Warning;
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

        private void ReviewTaskRequirements(string requirementType)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
            _cells.ClearTableRange(TableName03);
            var taskCells = new ExcelStyleCells(_excelApp, SheetName02);
            taskCells.SetFixedWorkingWorkSheet(true);

            var resultColumn = ResultColumn03;
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
                    var reqList = WorkOrderTaskActions.FetchRequirements(_eFunctions, d.DistrictCode, d.WorkOrder, requirementType, d.WoTaskNo);

                    var distinctReqList = reqList.GroupBy(x => new { x.DistrictCode, x.WorkGroup, x.WorkOrder, x.WoTaskNo, x.ReqCode}).Select(y => y.First());

                    foreach (var req in distinctReqList)
                    {
                        _cells.GetRange(1, i, resultColumn, i).ClearFormats();
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + req.DistrictCode; //DistrictCode
                        _cells.GetCell(2, i).Value = "" + req.WorkGroup;    //WorkGroup
                        _cells.GetCell(3, i).Value = "'" + req.WorkOrder;    //WorkOrder 
                        _cells.GetCell(4, i).Value = "'" + req.WoTaskNo;     //WoTaskNo 
                        _cells.GetCell(5, i).Value = "" + req.WoTaskDesc;   //WoTaskDesc 
                        _cells.GetCell(6, i).Value = "M";
                        _cells.GetCell(7, i).Value = "" + req.ReqType;      //ReqType 
                        _cells.GetCell(8, i).Value = "" + req.SeqNo;        //SeqNo 
                        _cells.GetCell(9, i).Value = "" + req.ReqCode;      //ReqCode
                        _cells.GetCell(10, i).Value = "" + req.ReqDesc;     //ReqDesc
                        _cells.GetCell(11, i).Value = "" + req.UoM;         //UoM
                        _cells.GetCell(12, i).Value = "" + req.EstSize;      //EstSize
                        _cells.GetCell(13, i).Value = "" + req.UnitsQty;      //UnitsQty
                        _cells.GetCell(14, i).Value = "" + req.RealQty;     //RealQty
                        _cells.GetCell(15, i).Value = "" + req.SharedTasks;     //SharedTask
                        if (Convert.ToInt16(req.SharedTasks) != 1)
                            _cells.GetRange(14, i, 15, i).Style = StyleConstants.Warning;
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
                    Debugger.LogError("RibbonEllipse.cs:ReviewTaskRequirements()", ex.Message);
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
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);


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
                        PlanStartTime = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value),
                        PlanFinishDate = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value),
                        PlanFinishTime = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value),
                        EstimatedDurationsHrs = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value),
                        NoLabor = _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value),
                        NoMaterial = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value),
                        AplEquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value),
                        AplType = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value),
                        AplCompCode = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value),
                        AplCompModCode = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value),
                        AplSeqNo = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value),
                        ExtTaskText = _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value),
                        CompletedCode = MyUtilities.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(27, i).Value)),
                        CompletedBy = _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value),
                        ClosedDate = _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value),
                        CompleteTaskText = _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value)
                    };

                    if (!string.IsNullOrWhiteSpace(woTask.PlanStartTime))
                        woTask.PlanFinishTime = woTask.PlanStartTime.PadLeft(6, '0');

                    if (!string.IsNullOrWhiteSpace(woTask.PlanFinishTime))
                        woTask.PlanFinishTime = woTask.PlanFinishTime.PadLeft(6, '0');
            
                    woTask.SetWorkOrderDto(woTask.WorkOrder);

                    if (string.IsNullOrWhiteSpace(action))
                        continue;

                    ReplyMessage replyMsg = null;

                    if (action.Equals(WorkOrderTaskActions.Modify))
                    {
                        var reply = WorkOrderTaskActions.ModifyWorkOrderTask(urlService, opSheet, woTask);
                        if (cbFlagEstDuration.Checked)
                        {
                            _cells.GetCell(14, i).Value = reply.planStrDate;
                            _cells.GetCell(15, i).Value = reply.planStrTime;
                            _cells.GetCell(16, i).Value = reply.planFinDate;
                            _cells.GetCell(17, i).Value = reply.planFinTime;
                        }
                        WorkOrderTaskActions.SetWorkOrderTaskText(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), _frmAuth.EllipseDstrct, _frmAuth.EllipsePost, true, woTask);
                    }
                    else if (action.Equals(WorkOrderTaskActions.Create))
                    {
                        var reply = WorkOrderTaskActions.CreateWorkOrderTask(urlService, opSheet, woTask);
                        if (cbFlagEstDuration.Checked)
                        {
                            _cells.GetCell(14, i).Value = reply.planStrDate;
                            _cells.GetCell(15, i).Value = reply.planStrTime;
                            _cells.GetCell(16, i).Value = reply.planFinDate;
                            _cells.GetCell(17, i).Value = reply.planFinTime;
                        }
                        WorkOrderTaskActions.SetWorkOrderTaskText(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), _frmAuth.EllipseDstrct, _frmAuth.EllipsePost, true, woTask);
                    }
                    else if (action.Equals(WorkOrderTaskActions.Delete))
                    {
                        var reply = WorkOrderTaskActions.DeleteWorkOrderTask(urlService, opSheet, woTask);
                    }
                    else if (action.Equals(WorkOrderTaskActions.Close))
                    {
                        replyMsg = WorkOrderTaskActions.CompleteWorkOrderTask(urlService, opSheet, woTask);
                        WorkOrderTaskActions.SetWorkOrderTaskText(Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label), _frmAuth.EllipseDstrct, _frmAuth.EllipsePost, true, woTask);
                    }
                    else if (action.Equals(WorkOrderTaskActions.ReOpen))
                    {
                        replyMsg = WorkOrderTaskActions.ReOpenWorkOrderTask(urlService, opSheet, woTask);
                    }
                    else
                        continue;

                    string messageResult = replyMsg == null ? "OK" : replyMsg.Message;

                    _cells.GetCell(ResultColumn02, i).Value = messageResult;
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
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
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
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName03, ResultColumn03);
            var i = TitleRow03 + 1;

            var opContextResource = new ResourceReqmntsService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };
            var opContextMaterial = new MaterialReqmntsService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };
            var opContextEquipment = new EquipmentReqmntsService.OperationContext()
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value) && !string.IsNullOrEmpty("" + _cells.GetCell(4, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var taskReq = new TaskRequirement();
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value);                         

                    taskReq.DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);                  
                    taskReq.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);                     
                    taskReq.WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);                     
                    taskReq.WoTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);                      
                    taskReq.WoTaskNo = string.IsNullOrWhiteSpace(taskReq.WoTaskNo) ? "001" : taskReq.WoTaskNo;
                    taskReq.WoTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);                    
                    taskReq.ReqType = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value);                       
                    taskReq.SeqNo = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value);                         
                    taskReq.ReqCode = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value);                       
                    taskReq.ReqDesc = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value);                      
                    taskReq.UoM = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value);                          
                    taskReq.EstSize = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value);                        
                    taskReq.UnitsQty = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value);                        
                    taskReq.RealQty = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value);                      


                    if (string.IsNullOrWhiteSpace(action))
                        continue;
                    else if (action.Equals("C"))
                    {
                        if (taskReq.ReqType.Equals(RequirementType.Labour.Key))
                            WorkOrderTaskActions.CreateTaskResource(urlService, opContextResource, taskReq);
                        else if (taskReq.ReqType.Equals(RequirementType.Material.Key))
                            WorkOrderTaskActions.CreateTaskMaterial(urlService, opContextMaterial, taskReq);
                        else if (taskReq.ReqType.Equals(RequirementType.Equipment.Key))
                            WorkOrderTaskActions.CreateTaskEquipment(urlService, opContextEquipment, taskReq);
                    }
                    else if (action.Equals("M"))
                    {
                        if (taskReq.ReqType.Equals(RequirementType.Labour.Key))
                            WorkOrderTaskActions.ModifyTaskResource(urlService, opContextResource, taskReq);
                        else if (taskReq.ReqType.Equals(RequirementType.Material.Key))
                            WorkOrderTaskActions.ModifyTaskMaterial(urlService, opContextMaterial, taskReq);
                        else if (taskReq.ReqType.Equals(RequirementType.Equipment.Key))
                            WorkOrderTaskActions.ModifyTaskEquipment(urlService, opContextEquipment, taskReq);
                    }
                    else if (action.Equals("D"))
                    {
                        if (taskReq.ReqType.Equals(RequirementType.Labour.Key))
                            WorkOrderTaskActions.DeleteTaskResource(urlService, opContextResource, taskReq);
                        else if (taskReq.ReqType.Equals(RequirementType.Material.Key))
                            WorkOrderTaskActions.DeleteTaskMaterial(urlService, opContextMaterial, taskReq);
                        else if (taskReq.ReqType.Equals(RequirementType.Equipment.Key))
                            WorkOrderTaskActions.DeleteTaskEquipment(urlService, opContextEquipment, taskReq);
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

        private void btnReviewMatRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewRequirements(RequirementType.Material.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnReviewEqpRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewRequirements(RequirementType.Equipment.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnReviewLabRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewRequirements(RequirementType.Labour.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnToDoReviewWorkOrders_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Esta opción está obsoleta para Ellipse 9 y estará disponible solamente hasta finalizado el proceso de migración. ¿Está seguro que desea continuar?", "Acción Obsoleta", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName08)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    _thread = new Thread(ReviewWoToDo);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoToDo()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnToDoReviewTasks_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Esta opción está obsoleta para Ellipse 9 y estará disponible solamente hasta finalizado el proceso de migración. ¿Está seguro que desea continuar?", "Acción Obsoleta", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName08)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    _thread = new Thread(ReviewWoTaskToDo);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoTaskToDo()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCleanTasksTable_Click(object sender, RibbonControlEventArgs e)
        {
            CleanTable(TableName02);
        }

        private void btnCleanToDo_Click(object sender, RibbonControlEventArgs e)
        {
            CleanTable(TableName08);
        }

        private void btnCleanRequirementTable_Click(object sender, RibbonControlEventArgs e)
        {
            CleanTable(TableName04);
        }

        private void CleanTable(string tableName)
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.ClearTableRange(tableName);
        }

        private void btnCreateToDo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Esta opción está obsoleta para Ellipse 9 y estará disponible solamente hasta finalizado el proceso de migración. ¿Está seguro que desea continuar?", "Acción Obsoleta", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName08)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(CreateWoToDoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateWoToDoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnDeleteToDo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Esta opción está obsoleta para Ellipse 9 y estará disponible solamente hasta finalizado el proceso de migración. ¿Está seguro que desea continuar?", "Acción Obsoleta", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName08)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(DeleteWoToDoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:DeleteWoToDoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateToDo_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (MessageBox.Show("Esta opción está obsoleta para Ellipse 9 y estará disponible solamente hasta finalizado el proceso de migración. ¿Está seguro que desea continuar?", "Acción Obsoleta", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return;
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName08)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(UpdateWoToDoList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateWoToDoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnReviewTaskRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewTaskRequirements(RequirementType.All.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnReviewTaskLabRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewTaskRequirements(RequirementType.Labour.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnReviewTaskMatRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewTaskRequirements(RequirementType.Material.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnReviewTaskEqpRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => ReviewTaskRequirements(RequirementType.Equipment.Key));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
    }

}
