﻿using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse.Connections;
using EllipseWorkOrdersClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Debugger = SharedClassLibrary.Utilities.Debugger;

namespace EllipseFinalizeWorkOrderExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const string SheetName01 = "WorkOrders";
        private const int TitleRow01 = 9;
        private const int ResultColumn01 = 15;
        private const string TableName01 = "WorkOrderTable";
        private const string ValidationSheetName = "ValidationSheetWorkOrder";
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();

        }
        public void LoadSettings()
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

            //settings.SetDefaultCustomSettingValue("OptionName1", "false");
            //settings.SetDefaultCustomSettingValue("OptionName2", "OptionValue2");
            //settings.SetDefaultCustomSettingValue("OptionName3", "OptionValue3");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //var optionItem1Value = MyUtilities.IsTrue(settings.GetCustomSettingValue("OptionName1"));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            //cbCustomSettingOption.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
        }
        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }
        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewWoList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnReReview_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReReviewWoList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnFinalize_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(FinalizeWoList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        
        private void btnCleanWorkOrderSheet_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
        }

        
        public void FormatSheet()
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
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "FINALIZE WORK ORDERS - ELLIPSE";
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
                var managementAreaList = ManagementArea.GetManagementArea().Select(g => g.Key).ToList();
                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();
                var dateCriteriaList = SearchDateCriteriaType.GetSearchDateCriteriaTypes().Select(g => g.Value).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = Districts.DefaultDistrict;
                _cells.SetValidationList(_cells.GetCell("B3"), districtList, ValidationSheetName, 1);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.Area.Value;
                _cells.SetValidationList(_cells.GetCell("A4"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), managementAreaList, ValidationSheetName, 3, false);
                _cells.SetValidationList(_cells.GetCell("B5"), workGroupList, ValidationSheetName, 4, false);
                _cells.GetCell("B4").Value = "MDC";
                _cells.GetCell("A5").Value = SearchFieldCriteriaType.WorkGroup.Value;
                _cells.SetValidationList(_cells.GetCell("A5"), ValidationSheetName, 2);
                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "FECHA";
                _cells.GetCell("D3").Value = SearchDateCriteriaType.NotFinalized.Value;
                _cells.SetValidationList(_cells.GetCell("D3"), dateCriteriaList, ValidationSheetName, 5);
                _cells.GetCell("C4").Value = "DESDE";
                _cells.GetCell("D4").Value = $"{DateTime.Now.Year:0000}" + "0101";
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetCell("C5").Value = "HASTA";
                _cells.GetCell("D5").Value = $"{DateTime.Now.Year:0000}" + $"{DateTime.Now.Month:00}" + $"{DateTime.Now.Day:00}";
                _cells.GetCell("D5").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);

                //
                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleInformation;
                
                //GENERAL
                _cells.GetCell(01, TitleRow01).Value = "DISTRICT";
                _cells.GetCell(01, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(02, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(02, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(03, TitleRow01).Value = "WORK_ORDER";
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(04, TitleRow01).Value = "WO_STATUS";
                _cells.GetCell(04, TitleRow01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(05, TitleRow01).Value = "DESCRIPTION";
                _cells.GetCell(06, TitleRow01).Value = "EQUIPMENT";
                _cells.GetCell(07, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(08, TitleRow01).Value = "MOD_CODE";
                _cells.GetCell(09, TitleRow01).Value = "RAISED_DATE";
                _cells.GetCell(09, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(10, TitleRow01).Value = "RAISED_TIME";
                _cells.GetCell(10, TitleRow01).AddComment("hhmmss");
                //COMPLETION INFO
                _cells.GetCell(11, TitleRow01).Value = "COMPL_COD";
                _cells.GetCell(11, TitleRow01).AddComment("Código de cierre de la orden");
                _cells.GetCell(12, TitleRow01).Value = "COMP_COMM";
                _cells.GetCell(12, TitleRow01).AddComment("Indica si una orden tiene comentarios de cierre");
                _cells.GetCell(13, TitleRow01).Value = "CLOSED DATE";
                _cells.GetCell(14, TitleRow01).Value = "COMPL_BY";
                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;
				_cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01+1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        
        public void ReviewWoList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRange(TableName01);

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


            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var listwo = WorkOrderActions.FetchWorkOrder(_eFunctions, district, searchCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, searchCriteriaValue2, dateCriteriaKey, startDate, endDate, statusKey);
            var i = TitleRow01 + 1;
            foreach (var wo in listwo)
            {
                try
                {
                    //Para resetear el estilo
                    _cells.GetRange(01, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(01, i).Value = "" + wo.districtCode;
                    _cells.GetCell(02, i).Value = "" + wo.workGroup;

                    _cells.GetCell(03, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(04, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(05, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(06, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(07, i).Value = "" + wo.compCode;
                    _cells.GetCell(08, i).Value = "" + wo.compModCode;
                    //if (wo.workOrderStatusM != "C" && string.IsNullOrEmpty(wo.workOrderStatusU) && !WorkOrderActions.ValidateUserStatus(wo.raisedDate, 60))
                    //    _cells.GetCell(10, i).Style = StyleConstants.Warning;
                    //else
                    //    _cells.GetCell(10, i).Style = StyleConstants.Normal;
                    
                    //DETAILS
                    _cells.GetCell(09, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(10, i).Value = "'" + wo.raisedTime;
                    _cells.GetCell(11, i).Value = "'" + wo.completedCode;
                    _cells.GetCell(12, i).Value = "" + wo.completeTextFlag;
                    if (wo.completeTextFlag == "N")
                        _cells.GetCell(12, i).Style = StyleConstants.Warning;
                    _cells.GetCell(13, i).Value = "" + wo.closeCommitDate;
                    _cells.GetCell(14, i).Value = "" + wo.completedBy;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", ex.Message);
                }
                finally
                {
					_cells.GetCell(1, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }
        
        public void ReReviewWoList()
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
                    var woNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value), woNo);
					
					if(wo == null || wo.GetWorkOrderDto().no == null)
                        throw new Exception ("WORK ORDER NO ENCONTRADA");
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(01, i).Value = "" + wo.districtCode;
                    _cells.GetCell(02, i).Value = "" + wo.workGroup;
                    _cells.GetCell(03, i).Value = "'" + wo.GetWorkOrderDto().prefix + wo.GetWorkOrderDto().no;
                    _cells.GetCell(04, i).Value = "" + WoStatusList.GetStatusName(wo.workOrderStatusM);
                    _cells.GetCell(05, i).Value = "" + wo.workOrderDesc;
                    _cells.GetCell(06, i).Value = "'" + wo.equipmentNo;
                    _cells.GetCell(07, i).Value = "" + wo.compCode;
                    _cells.GetCell(08, i).Value = "" + wo.compModCode;
                    //if (wo.workOrderStatusM != "C" && string.IsNullOrEmpty(wo.workOrderStatusU) &&
                    //    !WorkOrderActions.ValidateUserStatus(wo.raisedDate, 60))
                    //    _cells.GetCell(10, i).Style = StyleConstants.Warning;
                    //else
                    //    _cells.GetCell(10, i).Style = StyleConstants.Normal;
                    //DETAILS
                    _cells.GetCell(09, i).Value = "'" + wo.raisedDate;
                    _cells.GetCell(10, i).Value = "'" + wo.raisedTime;
                    _cells.GetCell(11, i).Value = "'" + wo.completedCode;
                    _cells.GetCell(12, i).Value = "" + wo.completeTextFlag;
                    if (wo.completeTextFlag == "N")
                        _cells.GetCell(12, i).Style = StyleConstants.Warning;
                    _cells.GetCell(13, i).Value = "'" + wo.closeCommitDate;
                    _cells.GetCell(14, i).Value = "'" + wo.completedBy;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewWOList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(1, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }
        public void FinalizeWoList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;

            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var environment = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var wo = new WorkOrder();
                    //GENERAL
                    wo.districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    wo.SetWorkOrderDto(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value));

                    var reply = WorkOrderActions.FinalizeWorkOrder(environment, opSheet, wo);
                    if (reply.finalCosts)
                    {
                        _cells.GetCell(ResultColumn01, i).Value = "FINALIZADA";
                        _cells.GetCell(1, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    }
                    else
                    {
                        _cells.GetCell(ResultColumn01, i).Value = "NO SE REALIZÓ ACCIÓN";
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:FinalizeWOList()", ex.Message);
                }
                finally
                {
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
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

        private void butAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

}
