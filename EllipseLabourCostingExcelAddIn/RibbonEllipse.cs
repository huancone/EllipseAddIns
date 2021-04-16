using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Utilities;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using System.Web.Services.Ellipse;
using EllipseWorkOrdersClassLibrary;
using EllipseStdTextClassLibrary;
using System.Threading;
using EllipseLabourCostingExcelAddIn.EquipmentHireLibrary;
using EllipseLabourCostingExcelAddIn.LabourEmployeeLibrary;

using OperationContext = EllipseLabourCostingExcelAddIn.LabourCostingTransService.OperationContext;
using EquipmentHireOperationContext = EllipseLabourCostingExcelAddIn.EquipHireTranService.OperationContext;

namespace EllipseLabourCostingExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        Excel.Application _excelApp;
        const string SheetName01 = "Labour";
        private const string ValidationSheetName = "ValidationSheetLabour";

        const int DefaultStartCol = 3;
        const int NroEle = 3;//cantidad de elementos por orden a ingresar (HRS, TSK, Earning Code)

        const int DefaultTitleRow = 9;
        const int Mso850TitleRow = 5;//Deprecated
        const int Mse850TitleRow = 5;
        const int ElecsaTitleRow = 6;
        const int EquipmentHireTitleRow = 5;
        
        private const int ElecsaTitleColumn = 6;
        private const int Mso850ResultColumn = 18;//Deprecated
        private const int Mse850ResultColumn = 14;
        private const int ElecsaResultColumn = 30;
        private const int EquipmentHireResultColumn = 11;
        
        private const string TableNameMso850 = "Mso850Table";//Deprecated
        private const string TableNameMse850 = "Mse850Table";
        private const string TableNameDefault = "LabourDefaultTable";
        private const string TableNameElecsa = "ElecsaTable";
        private const string TableNameEquipmentHire = "EquipmentHireTable";
        private const int OtFields = 20;//Para Group Default

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

            /* Example Configuration Settings
            settings.SetDefaultCustomSettingValue("AutoSort", "Y");
            settings.SetDefaultCustomSettingValue("OverrideAccountCode", "Maintenance");
            settings.SetDefaultCustomSettingValue("IgnoreItemError", "N");
            
            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {
            
                MessageBox.Show(ex.Message, "Load Settings", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            
            var overrideAccountCode = settings.GetCustomSettingValue("OverrideAccountCode");
            if (overrideAccountCode.Equals("Maintenance"))
                cbAccountElementOverrideMntto.Checked = true;
            else if (overrideAccountCode.Equals("Disable"))
                cbAccountElementOverrideDisable.Checked = true;
            else if (overrideAccountCode.Equals("Alwats"))
                cbAccountElementOverrideAlways.Checked = true;
            else if (overrideAccountCode.Equals("Default"))
                cbAccountElementOverrideDefault.Checked = true;
            else
                cbAccountElementOverrideDefault.Checked = true;
            cbAutoSortItems.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("AutoSort"));
            cbIgnoreItemError.Checked = MyUtilities.IsTrue(settings.GetCustomSettingValue("IgnoreItemError"));


            settings.SaveCustomSettings();
            */
        }

        private void btnFormatHeader_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetHeaderData();
        }

        private void btnFormatGroupEmployee_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Contains(SheetName01))
            {
                var groupName = "" + _cells.GetCell("B4").Value;

                if (!groupName.Equals(""))
                    FormatGroupEmployees(groupName);
                else
                    MessageBox.Show(@"No se ha seleccionado ningún grupo");
            }
            else
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para el cargue de labor");
        }

        private void btnFormatMso850_Click(object sender, RibbonControlEventArgs e)
        {
            FormatMse850Sheet();
        }

        private void btnFormatElecsa_Click(object sender, RibbonControlEventArgs e)
        {
            FormatElecsaSheet();
        }

        private void btnFormatEquipmentLabour_Click(object sender, RibbonControlEventArgs e)
        {
            FormatEquipmentHire();
        }
        private void btnLoadLaborSheet_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                if (!_cells.IsDecimalDotSeparator())
                    if (MessageBox.Show(@"El separador de decimales configurado actualmente no es el punto. Usar un separador de decimales diferente puede generar errores al momento de cargar valores numéricos. ¿Está seguro que desea continuar?", @"ALERTA DE SEPARADOR DE DECIMALES", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Default))
                    _thread = new Thread(LoadDefaultLabourCost);
                //else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mso850)) //Deprecated
                //    _thread = new Thread(LoadMso850LabourCost);
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mse850))
                    _thread = new Thread(LoadMse850LabourCost);
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Elecsa))
                    _thread = new Thread(LoadElecsaLabourCost);
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.EquipmentHire))
                    _thread = new Thread(LoadEquipmentHire);
                else
                {
                    MessageBox.Show(@"La hoja de Excel no tiene el formato válido para el cargue de labor");
                    return;
                }
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:LoadLaborSheet()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCleanSheet_Click(object sender, RibbonControlEventArgs e)
        {
            CleanLabourTable();
        }

        public void FormatSheetHeaderData()
        {
            try
            {
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "LOAD GROUP EMPLOYEES LABOUR - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

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

                _cells.GetCell("A3").Value = "DISTRICT";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "WORKGROUP";
                _cells.GetCell("A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("A5").Value = "FECHA";
                _cells.GetCell("A5").AddComment("YYYYMMDD");
                _cells.GetCell("A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B5").Style = _cells.GetStyle(StyleConstants.Select);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        public void FormatGroupEmployees(string groupName)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01 + LabourSheetTypeConstants.Default;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetRange("A6", "AZ65536").EntireRow.Delete(Excel.XlDirection.xlUp);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var titleRow = DefaultTitleRow;
                //var resultColumn = 0;
                var tableName = TableNameDefault;

                var sqlQuery = Queries.GetGroupEmployeesQuery(groupName, _eFunctions.DbReference, _eFunctions.DbLink);
                var drEmployees = _eFunctions.GetQueryResult(sqlQuery);

                _cells.GetCell("A6").Value = "NÚMERO OT";
                _cells.GetCell("A6").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("A6").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                _cells.GetCell("A6").Borders.Weight = 2d;
                _cells.MergeCells("A6", "B7");
                _cells.GetRange("A6", "B7").BorderAround2(Type.Missing, Excel.XlBorderWeight.xlMedium);

                _cells.GetCell("A8").Value = "DESCRIPCIÓN OT";
                _cells.GetCell("A8").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.MergeCells("A8", "B8");
                _cells.GetRange("A8", "B8").BorderAround2(Type.Missing, Excel.XlBorderWeight.xlMedium);

                _cells.GetCell("A9").Value = "EMPLEADO";
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("B9").Value = "CÉDULA";
                _cells.GetCell("B9").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                Worksheet vstoSheet =
                    Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
                for (var i = 0; i < OtFields; i++)
                {
                    //Título Nro OT
                    _cells.GetCell(DefaultStartCol + i * NroEle, titleRow - 3).Value = "OT " + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle, titleRow - 3).Style =
                        _cells.GetStyle(StyleConstants.TitleOptional);
                    _cells.MergeCells(DefaultStartCol + i * NroEle, titleRow - 3,
                        DefaultStartCol + i * NroEle + (NroEle - 1), titleRow - 3);
                    //Número OT
                    _cells.GetCell(DefaultStartCol + i * NroEle, titleRow - 2).Style =
                        _cells.GetStyle(StyleConstants.Select);
                    vstoSheet.Controls.Remove("SeekOrder" + i);
                    var orderNameRange =
                        vstoSheet.Controls.AddNamedRange(
                            _cells.GetCell(DefaultStartCol + i * NroEle, titleRow - 2),
                            "SeekOrder" + i);
                    orderNameRange.Change += GetDefaultWorkOrderDescriptionChangedValue;
                    _cells.GetRange(DefaultStartCol + i * NroEle, titleRow - 2,
                        DefaultStartCol + i * NroEle + (NroEle - 1), titleRow - 2)
                        .NumberFormat = NumberFormatConstants.Text;
                    _cells.MergeCells(DefaultStartCol + i * NroEle, titleRow - 2,
                        DefaultStartCol + i * NroEle + (NroEle - 1), titleRow - 2);
                    //Descripción OT
                    _cells.GetCell(DefaultStartCol + i * NroEle, titleRow - 1).Style =
                        _cells.GetStyle(StyleConstants.Option);
                    _cells.MergeCells(DefaultStartCol + i * NroEle, titleRow - 1,
                        DefaultStartCol + i * NroEle + (NroEle - 1), titleRow - 1);
                    //Cada componente
                    _cells.GetCell(DefaultStartCol + i * NroEle + 0, titleRow).Value = "HRS_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 0, titleRow).AddComment("hh.mm");
                    _cells.GetCell(DefaultStartCol + i * NroEle + 0, titleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleRequired);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 1, titleRow).Value = "TASK_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 1, titleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleRequired);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 2, titleRow).Value = "ECODE_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 2, titleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleOptional);

                }
                _cells.GetRange(DefaultStartCol, titleRow + 1, DefaultStartCol + OtFields * NroEle,
                    titleRow + 1)
                    .NumberFormat = NumberFormatConstants.Text;
                //Validación para los Earning Codes
                var itemList = _eFunctions.GetItemCodes("EA");
                var validationList = itemList.Select(item => item.Code + " - " + item.Description).ToList();

                //creo la validación y la asigno al primer elemento
                _cells.SetValidationList(_cells.GetCell(DefaultStartCol + 2, titleRow + 1), validationList,
                    ValidationSheetName, 1);
                var validation = _cells.GetCell(DefaultStartCol + 2, titleRow + 1).Validation;
                //asigno la validación al resto de los elementos
                for (var j = 1; j < OtFields; j++)
                {
                    _cells.GetCell(DefaultStartCol + j * NroEle + 2, titleRow + 1).Validation.Delete();
                    _cells.GetCell(DefaultStartCol + j * NroEle + 2, titleRow + 1)
                        .Validation.Add((Excel.XlDVType)validation.Type, validation.AlertStyle, validation.Operator,
                            validation.Formula1, validation.Formula2);
                }
                _cells.FormatAsTable(
                    _cells.GetRange(1, titleRow, DefaultStartCol + OtFields * NroEle - 1, titleRow + 1),
                    tableName);
                var lengthEmployees = 0;
                if (drEmployees != null && !drEmployees.IsClosed)
                {
                    while (drEmployees.Read())
                    {
                        _cells.GetCell(1, titleRow + 1 + lengthEmployees).Value = drEmployees["NOMBRE"].ToString().Trim();
                        _cells.GetCell(2, titleRow + 1 + lengthEmployees).Value = "'" + drEmployees["CEDULA"].ToString().Trim();
                        lengthEmployees++;
                    }
                }
                else
                    MessageBox.Show(@"No se han encontrado datos para el modelo especificado");

                
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatGroupDefault()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar obtener el formato del grupo seleccionado. " + ex.Message);
            }
            finally
            {
				_eFunctions.CloseConnection();
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void FormatMse850Sheet()
        {
            try
            {
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01 + LabourSheetTypeConstants.Mse850;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var titleRow = Mse850TitleRow;
                var resultColumn = Mse850ResultColumn;
                var tableName = TableNameMse850;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "LOAD EMPLOYEES LABOUR - ELLIPSE 8";
                _cells.GetCell("B1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;

                //listas de validación
                //var itemListEc = _eFunctions.GetItemCodes("EA");//earning codes
                //var earningCodeList = itemListEc.Select(item => item.code + " - " + item.description).ToList();

                //var itemListLc = _eFunctions.GetItemCodes("LC");//laborclass codes
                //var laborClassList = itemListLc.Select(item => item.code + " - " + item.description).ToList();

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetRange(1, titleRow + 1, resultColumn - 1, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(1, titleRow).Value = "Trans.Date";
                _cells.GetCell(1, titleRow).AddComment("yyyyMMss");
                _cells.GetCell(2, titleRow).Value = "EmployeeId";
                _cells.GetCell(3, titleRow).Value = "InterDistrict";
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, titleRow).Value = "Project";
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, titleRow).Value = "WorkOrder";
                _cells.GetCell(6, titleRow).Value = "WOTask";
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow).Value = "EquipmentRef";
                _cells.GetCell(8, titleRow).Value = "EquipmentNo";
                _cells.GetCell(8, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(9, titleRow).Value = "LaborClass";
                //_cells.SetValidationList(_cells.GetCell(9, titleRow + 1), laborClassList, ValidationSheetName, 2);
                _cells.GetCell(9, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, titleRow).Value = "EarningCode";
                //_cells.SetValidationList(_cells.GetCell(10, titleRow + 1), earningCodeList, ValidationSheetName, 3);
                _cells.GetCell(10, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, titleRow).Value = "Hours";
                _cells.GetCell(11, titleRow).AddComment("hh.mm");
                _cells.GetCell(12, titleRow).Value = "OvertimeIndicator";
                _cells.GetCell(12, titleRow).AddComment("Y/N");
                _cells.GetCell(12, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(13, titleRow).Value = "AccountCode";
                _cells.GetCell(13, titleRow).Style = StyleConstants.TitleOptional;

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatMse850Sheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void FormatMso850Sheet()
        {
            try
            {
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01 + LabourSheetTypeConstants.Mso850;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var titleRow = Mso850TitleRow;
                var resultColumn = Mso850ResultColumn;
                var tableName = TableNameMso850;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "LOAD EMPLOYEES LABOUR - ELLIPSE 8";
                _cells.GetCell("B1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;

                //listas de validación
                var itemList1 = _eFunctions.GetItemCodes("SC");//comp codes
                var compCodeList = itemList1.Select(item => item.Code + " - " + item.Description).ToList();

                var itemList2 = _eFunctions.GetItemCodes("EA");//earning codes
                var earningCodeList = itemList2.Select(item => item.Code + " - " + item.Description).ToList();

                var itemList3 = _eFunctions.GetItemCodes("LC");//laborclass codes
                var laborClassList = itemList3.Select(item => item.Code + " - " + item.Description).ToList();

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetRange(1, titleRow + 1, resultColumn - 1, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(1, titleRow).Value = "TRAN_DATE";
                _cells.GetCell(1, titleRow).AddComment("yyyyMMss");
                _cells.GetCell(2, titleRow).Value = "ORD_HRS";
                _cells.GetCell(2, titleRow).AddComment("hh.mm");
                _cells.GetCell(3, titleRow).Value = "OT_HRS";
                _cells.GetCell(3, titleRow).AddComment("hh.mm");
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, titleRow).Value = "VALUE";
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, titleRow + 1).Style = StyleConstants.Disabled;
                _cells.GetCell(5, titleRow).Value = "INT_DSTRCT_CDE";
                _cells.GetCell(6, titleRow).Value = "ACCOUNT";
                _cells.GetCell(7, titleRow).Value = "STATUS";
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow + 1).Style = StyleConstants.Disabled;
                _cells.GetCell(8, titleRow).Value = "WOP_IND";
                _cells.GetCell(8, titleRow).AddComment("P: Proyecto \nW: WorkOrder");
                _cells.SetValidationList(_cells.GetCell(8, titleRow + 1), new List<string> { "P", "W" }, ValidationSheetName, 1);
                _cells.GetCell(9, titleRow).Value = "WO_PROJ";
                _cells.GetCell(10, titleRow).Value = "TASK";
                _cells.GetCell(11, titleRow).Value = "EMP_ID";
                _cells.GetCell(12, titleRow).Value = "EQUIPMENT";
                _cells.GetCell(13, titleRow).Value = "UNITS_COMP";
                _cells.GetCell(14, titleRow).Value = "PC_COMP";
                _cells.GetCell(15, titleRow).Value = "CODE_COMP";
                _cells.SetValidationList(_cells.GetCell(15, titleRow + 1), compCodeList, ValidationSheetName, 2);
                _cells.GetCell(16, titleRow).Value = "EARN_CLASS";
                _cells.SetValidationList(_cells.GetCell(16, titleRow + 1), earningCodeList, ValidationSheetName, 3);
                _cells.GetCell(17, titleRow).Value = "LABOUR_CLASS";
                _cells.SetValidationList(_cells.GetCell(17, titleRow + 1), laborClassList, ValidationSheetName, 4);


                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.FormatAsTable(
                    _cells.GetRange(1, titleRow, resultColumn, titleRow + 1),
                    tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatMso850Sheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void FormatElecsaSheet()
        {
            try
            {
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01 + LabourSheetTypeConstants.Elecsa;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var titleRow = ElecsaTitleRow;
                var titleColumn = ElecsaTitleRow;
                var resultColumn = ElecsaResultColumn;
                var tableName = TableNameEquipmentHire;


                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "ELECSA SÁBANA DE LABOR - ELLIPSE 8";
                _cells.GetCell("B1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells(2, 1, resultColumn - 1, 2);

                _cells.GetCell(resultColumn, 1).Value = "OBLIGATORIO";
                _cells.GetCell(resultColumn, 1).Style = StyleConstants.TitleRequired;
                _cells.GetCell(resultColumn, 2).Value = "OPCIONAL";
                _cells.GetCell(resultColumn, 2).Style = StyleConstants.TitleOptional;
                _cells.GetCell(resultColumn, 3).Value = "INFORMATIVO";
                _cells.GetCell(resultColumn, 3).Style = StyleConstants.TitleInformation;

                //listas de validación
                //laborclass codes - descr
                var laborClassList = new List<string>
                {
                    "ELEU - ELECTRICISTA COSTO UNITARIO",
                    "ELEG - ELECTRICISTA COSTO GLOBAL",
                    "ELCU - COMUNICACIONES COSTO UNITARIO",
                    "ELCG - COMUNICACIONES COSTO GLOBAL",
                    "ELLI - ELECTRICISTA LINIEROS",
                    "ELAX - ELECTRICISTA AUXILIARES",
                    "ELTE - TECNÓLOGO ELECTRÓNICO",
                    "ELCS - COORDINADOR SIIO",
                    "ELIU - INGENIERO ELECTRICISTA"
                };

                _cells.GetCell(1, titleRow - 3).Value = "DISTRITO";
                _cells.GetCell(1, titleRow - 3).Style = StyleConstants.Option;
                _cells.GetCell(2, titleRow - 3).Value = "ICOR";
                _cells.GetCell(2, titleRow - 3).Style = StyleConstants.Select;
                _cells.GetCell(1, titleRow - 2).Value = "FECHA";
                _cells.GetCell(1, titleRow - 2).Style = StyleConstants.Option;
                _cells.GetCell(1, titleRow - 2).AddComment("YYYYMMDD");
                _cells.GetCell(2, titleRow - 2).Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell(2, titleRow - 2).Style = StyleConstants.Select;

                _cells.GetCell(3, titleRow - 3).Value = "TIPO CARGA";
                _cells.GetCell(3, titleRow - 3).Style = StyleConstants.Option;
                _cells.GetCell(4, titleRow - 3).Value = "LABOR Y DURACIÓN";
                _cells.GetCell(4, titleRow - 3).Style = StyleConstants.Select;
                var loadTypeList = new List<string>
                {
                    "LABOR",
                    "DURACIÓN",
                    "LABOR Y DURACIÓN"
                };

                _cells.GetCell(titleColumn - 1, titleRow - 2).Value = "EMPLEADOS";
                _cells.GetCell(titleColumn - 1, titleRow - 2).Style = StyleConstants.Option;
                _cells.GetCell(1, titleRow).Value = "CENTRO";
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(2, titleRow).Value = "WORKORDER";
                _cells.GetCell(2, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(3, titleRow).Value = "TASK";
                _cells.GetCell(4, titleRow).Value = "DESCRIPCIÓN";
                _cells.GetRange(2, titleRow, 4, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(titleColumn - 1, titleRow - 1).Value = "LABOR";
                _cells.GetCell(titleColumn - 1, titleRow - 1).Style = StyleConstants.Option;
                _cells.GetCell(titleColumn - 1, titleRow).Value = "Nro.";
                _cells.GetCell(titleColumn - 1, titleRow + 1).Style = StyleConstants.Disabled;
                for (var i = titleColumn; i < resultColumn; i++)
                {
                    _cells.GetCell(i, titleRow).Value = "" + (i - titleColumn + 1);
                    _cells.GetCell(i, titleRow).AddComment("hh.mm");
                }
                _cells.GetRange(titleColumn, titleRow, resultColumn - 1, titleRow).Style =
                    StyleConstants.TitleOptional;


                _cells.GetRange(titleColumn, titleRow - 2, resultColumn - 1, titleRow - 1).Style = StyleConstants.Select;
                //validaciones de campo
                _cells.SetValidationList(_cells.GetCell(2, titleRow - 3), Districts.GetDistrictList(), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell(4, titleRow - 3), loadTypeList, ValidationSheetName, 2);

                _cells.SetValidationList(_cells.GetRange(titleColumn, titleRow - 1, resultColumn - 1, titleRow - 1), laborClassList, ValidationSheetName, 3);
                _cells.GetRange(titleColumn - 1, titleRow - 2, resultColumn - 1, titleRow - 1).ColumnWidth = 3.57;
                _cells.GetRange(titleColumn - 1, titleRow - 2, resultColumn - 1, titleRow - 1).Orientation = 90;

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;

                _cells.GetCell(resultColumn + 1, titleRow).Value = "DURATION CODE";
                _cells.GetCell(resultColumn + 2, titleRow).Value = "HR INICIAL";
                _cells.GetCell(resultColumn + 2, titleRow).AddComment("hhmmss");
                _cells.GetCell(resultColumn + 2, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(resultColumn + 3, titleRow).Value = "HR FINAL";
                _cells.GetCell(resultColumn + 3, titleRow).AddComment("hhmmss");
                _cells.GetCell(resultColumn + 3, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(resultColumn + 4, titleRow).Value = "COMENTARIO";
                _cells.GetCell(resultColumn + 4, titleRow).AddComment("No modifica el comentario de cierre, solo adiciona al comentario existente");
                _cells.GetCell(resultColumn + 4, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetRange(resultColumn + 1, titleRow, resultColumn + 4, titleRow).Style = StyleConstants.TitleRequired;


                _cells.GetCell(resultColumn + 5, titleRow).Value = "RESULTADO DUR.";
                _cells.GetCell(resultColumn + 5, titleRow).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, titleRow + 1, resultColumn + 5, titleRow + 1).NumberFormat =
                    NumberFormatConstants.Text;

                //búsquedas especiales de tabla
                var table = _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn + 5, titleRow + 1), tableName);
                var tableObject = Globals.Factory.GetVstoObject(table);
                tableObject.Change += GetTableChangedValue;

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _cells.SetWorksheetVisibility(ValidationSheetName, false);

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatElecsaSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void FormatEquipmentHire()
        {
            try
            {
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01 + LabourSheetTypeConstants.EquipmentHire;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var titleRow = EquipmentHireTitleRow;
                var resultColumn = EquipmentHireResultColumn;
                var tableName = TableNameEquipmentHire;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "LOAD EQUIPMENT HIRE TRANSACTIONS MSO496 - ELLIPSE 8";
                _cells.GetCell("B1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;

                _cells.GetRange(1, titleRow, resultColumn - 1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetRange(1, titleRow + 1, resultColumn - 1, titleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(1, titleRow).Value = "Fecha Transacción";
                _cells.GetCell(1, titleRow).AddComment("yyyyMMss");
                _cells.GetCell(2, titleRow).Value = "Id Empleado";
                _cells.GetCell(3, titleRow).Value = "Secuencia";
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, titleRow).Value = "Indicador";
                _cells.GetCell(4, titleRow).AddComment("W: Orden de Trabajo (Predeterminada), P: Proyecto");
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, titleRow).Value = "Orden / Proyecto";
                _cells.GetCell(6, titleRow).Value = "Tarea";
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow).Value = "Centro Costo";
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(8, titleRow).Value = "Equipo";
                _cells.GetCell(9, titleRow).Value = "Estadística";
                _cells.GetCell(10, titleRow).Value = "Valor";

                var woProjectList = new List<string>();
                woProjectList.Add("W - WorkOrder");
                woProjectList.Add("P - Project");
                _cells.SetValidationList(_cells.GetCell(4, titleRow + 1), woProjectList, false);

                _cells.GetCell(resultColumn, titleRow).Value = "RESULTADO";
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatEquipmentHire()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void LoadDefaultLabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = DefaultTitleRow;
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                //Deprecated
                //var opSheet = new Screen.OperationContext
                //{
                //    district = _frmAuth.EllipseDstrct,
                //    position = _frmAuth.EllipsePost,
                //    maxInstances = 100,
                //    returnWarnings = Debugger.DebugWarnings
                //};
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = 10;
                //recorro la lista de empleados de forma vertical
                while ("" + _cells.GetCell(2, i).Value != "")
                {
                    //proceso por empleado
                    var employee = new LabourEmployee
                    {
                        Employee = ("" + _cells.GetCell(2, i).Value).Trim(),
                        TransactionDate = ("" + _cells.GetCell(2, 5).Value).Trim()
                    };

                    //recorro la lista de órdenes de forma horizontal
                    var j = DefaultStartCol;

                    while ("" + _cells.GetCell(j, titleRow - 2).Value != "")
                    {
                        try
                        {
                            if ("" + _cells.GetCell(j, i).Value != "" && "" + _cells.GetCell(j, i).Value != "0")
                            {
                                var earningClass = _cells.GetEmptyIfNull(_cells.GetCell(j + 2, i).Value2);
                                //obtengo solo el código sin la descripción
                                if (!earningClass.Equals("") && earningClass.Contains(" - "))
                                    earningClass = earningClass.Substring(0,
                                        earningClass.IndexOf(" - ", StringComparison.Ordinal));


                                employee.LabourCostingHours = _cells.GetEmptyIfNull(_cells.GetCell(j, i).Value2);
                                employee.WorkOrderTask = _cells.GetEmptyIfNull(_cells.GetCell(j + 1, i).Value2);
                                if (string.IsNullOrWhiteSpace(employee.WorkOrderTask) && cbAutoTaskAssigment.Checked)
                                    employee.WorkOrderTask = "001";
                                //el número de tarea debe tener tres dígitos 001
                                if (string.IsNullOrWhiteSpace(employee.WorkOrderTask) && employee.WorkOrderTask.Length >= 1 && employee.WorkOrderTask.Length < 3)
                                    employee.WorkOrderTask = employee.WorkOrderTask.PadLeft(3, '0');
                                employee.EarnCode = earningClass;

                                employee.WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(j, titleRow - 2).Value2);

                                LabourEmployeeActions.CreateEmployeeMse(urlService, opSheet, employee, cbReplaceExisting.Checked);

                                _cells.GetRange(j, i, j, i).ClearComments();
                                _cells.GetRange(j, i, j + NroEle - 1, i).Style = StyleConstants.Success;
                                _cells.GetRange(j, i, j, i).Select();
                            }
                        }
                        catch (Exception ex)
                        {
                            Debugger.LogError("RibbonEllipse:LoadDefaultLabourCost()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                            _cells.GetRange(j, i, j + NroEle - 1, i).Style = StyleConstants.Error;
                            _cells.GetRange(j, i, j, i).ClearComments();
                            _cells.GetRange(j, i, j, i).AddComment(ex.Message);
                            _cells.GetRange(j, i, j, i).Select();
                        }
                        finally
                        {
                            j = j + NroEle;
                        }
                    } //--while de órdenes
                    i++;
                } //--while de empleados
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadDefaultLabourCost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }

        }

        public void LoadMse850LabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = Mse850TitleRow;
                var resultColumn = Mse850ResultColumn;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = titleRow + 1;
                //recorro la lista de empleados de forma vertical
                while (!_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2).Equals(""))
                {
                    try
                    {
                        //valores de listas
                        var transactionDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                        var employeeId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                        var interDistrict = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                        var project = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                        var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                        var woTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                        var equipmentRef = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                        var equipmentNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value);
                        var laborClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value);
                        var earningCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                        var hours = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                        var overtimeIndicator = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);
                        var accountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value);

                        //obtengo solo el código sin la descripción
                        if (earningCode != null && earningCode.Contains(" - "))
                            earningCode = earningCode.Substring(0, earningCode.IndexOf(" - ", StringComparison.Ordinal));
                        if (laborClass != null && laborClass.Contains(" - "))
                            laborClass = laborClass.Substring(0, laborClass.IndexOf(" - ", StringComparison.Ordinal));

                        //proceso por empleado
                        var employee = new LabourEmployee
                        {
                            TransactionDate = transactionDate,
                            Employee = employeeId,
                            InterDistrictCode = interDistrict,
                            Project = project,
                            WorkOrder = workOrder,
                            WorkOrderTask = woTask,
                            EquipmentRef = equipmentRef,
                            EquipmentNo = equipmentNo,
                            LabourClass = laborClass,
                            EarnCode = earningCode,
                            LabourCostingHours = hours,
                            OvertimeInd = MyUtilities.IsTrue(overtimeIndicator),
                            AccountCode = accountCode,
                        };

//                        if (!string.IsNullOrWhiteSpace(employee.WorkOrder))
//                        {
//                            if (string.IsNullOrWhiteSpace(employee.WorkOrderTask) && cbAutoTaskAssigment.Checked)
//                                employee.WorkOrderTask = "001";
//                            if (!string.IsNullOrWhiteSpace(employee.WorkOrderTask) && employee.WorkOrderTask.Length >= 1 && employee.WorkOrderTask.Length < 3)
//                                employee.WorkOrderTask = employee.WorkOrderTask.PadLeft(3, '0');
//                        }

                        var reply = LabourEmployeeActions.CreateEmployeeMse(urlService, opSheet, employee, cbReplaceExisting.Checked);

                        if (reply.errors.Length == 0)
                        {
                            _cells.GetCell(resultColumn, i).Value = "SUCCESS";
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn, i).Select();
                        }
                        else
                        {
                            _cells.GetCell(resultColumn, i).Value = string.Join(",", reply.errors.Select(p => p.messageText).ToArray());
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                            _cells.GetCell(resultColumn, i).Select();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadMse850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    finally
                    {
                        i++;
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadMse850LabourCost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        //Deprecated
        [Obsolete("Not used anymore", true)]
        public void LoadMso850LabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = Mso850TitleRow;
                var resultColumn = Mso850ResultColumn;

                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = titleRow + 1;
                //recorro la lista de empleados de forma vertical
                while (!_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2).Equals(""))
                {
                    try
                    {
                        //valores de listas
                        var codeComp = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value);
                        var earningCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value);
                        var laborClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value);

                        //obtengo solo el código sin la descripción
                        if (codeComp != null && codeComp.Contains(" - "))
                            codeComp = codeComp.Substring(0, codeComp.IndexOf(" - ", StringComparison.Ordinal));
                        if (earningCode != null && earningCode.Contains(" - "))
                            earningCode = earningCode.Substring(0, earningCode.IndexOf(" - ", StringComparison.Ordinal));
                        if (laborClass != null && laborClass.Contains(" - "))
                            laborClass = laborClass.Substring(0, laborClass.IndexOf(" - ", StringComparison.Ordinal));

                        //proceso por empleado
                        var employee = new LabourEmployee
                        {
                            TransactionDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                            LabourCostingHours = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value),
                            OvertimeInd = MyUtilities.IsTrue(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value)),
                            LabourCostingValue = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value),
                            InterDistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                            AccountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value),
                            PostingStatus = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value),
                            Project = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                            WorkOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value),
                            WorkOrderTask = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                            Employee = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value),
                            EquipmentNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value),
                            UnitsComplete = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value),
                            PercentComplete = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value),
                            CompletedCode = codeComp,
                            EarnCode = earningCode,
                            LabourClass = laborClass
                        };

                        if (!string.IsNullOrWhiteSpace(employee.WorkOrder))
                        {
                            if (string.IsNullOrWhiteSpace(employee.WorkOrderTask) && cbAutoTaskAssigment.Checked)
                                employee.WorkOrderTask = "001";
                            if (!string.IsNullOrWhiteSpace(employee.WorkOrderTask) && employee.WorkOrderTask.Length >= 1 && employee.WorkOrderTask.Length < 3)
                                employee.WorkOrderTask = employee.WorkOrderTask.PadLeft(3, '0');
                        }


                        LabourEmployeeActions.CreateEmployeeMso(_eFunctions, urlService, opSheet, employee, cbAutoTaskAssigment.Checked, cbReplaceExisting.Checked);

                        _cells.GetCell(resultColumn, i).Value = "SUCCESS";
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    finally
                    {
                        i++;
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        public void LoadEquipmentHire()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = EquipmentHireTitleRow;
                var resultColumn = EquipmentHireResultColumn;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new EquipmentHireOperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = titleRow + 1;
                //recorro la lista de forma vertical
                while (!string.IsNullOrWhiteSpace(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2)))
                {
                    try
                    {
                        //preparación de valores
                        var indicator = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value));
                        indicator = string.IsNullOrWhiteSpace(indicator) ? "W" : indicator;
                        string workOrder = null;
                        string project = null;
                        if (string.IsNullOrWhiteSpace(indicator) || indicator.Equals("W"))
                            workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                        else
                            project = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);

                        //equipment hire
                        var equipmentHire = new EquipmentHire
                        {
                            TransactionDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                            EmployeeId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value),
                            TranSequence = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value),
                            WoProjectIndicator = indicator,
                            ProjectNo = project,
                            WorkOrder = workOrder,
                            Task = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value),
                            AccountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value),
                            EquipmentReference = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                            StatisticType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value),
                            Value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value),
                        };
                    

                        if (equipmentHire.WoProjectIndicator.Equals("W") && !string.IsNullOrWhiteSpace(equipmentHire.WorkOrder))
                        {
                            if (string.IsNullOrWhiteSpace(equipmentHire.Task) && cbAutoTaskAssigment.Checked)
                                equipmentHire.Task = "001";
                            else if (!string.IsNullOrWhiteSpace(equipmentHire.Task) && equipmentHire.Task.Length >= 1 && equipmentHire.Task.Length < 3)
                                equipmentHire.Task = equipmentHire.Task.PadLeft(3, '0');
                        }


                        EquipmentHireActions.CreateEquipmentHire(urlService, opSheet, equipmentHire);

                        _cells.GetCell(resultColumn, i).Value = "CREADO";
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadEquipmentHire()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    finally
                    {
                        i++;
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadEquipmentHire()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void DeleteEquipmentHire()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = EquipmentHireTitleRow;
                var resultColumn = EquipmentHireResultColumn;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new EquipmentHireOperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = titleRow + 1;
                //recorro la lista de forma vertical
                while (!string.IsNullOrWhiteSpace(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2)))
                {
                    try
                    {
                        //preparación de valores
                        var indicator = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value));
                        indicator = string.IsNullOrWhiteSpace(indicator) ? "W" : indicator;
                        string workOrder = null;
                        string project = null;
                        if (string.IsNullOrWhiteSpace(indicator) || indicator.Equals("W"))
                            workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                        else
                            project = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);

                        //equipment hire
                        var equipmentHire = new EquipmentHire
                        {
                            TransactionDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                            EmployeeId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value),
                            TranSequence = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value),
                            WoProjectIndicator = indicator,
                            ProjectNo = project,
                            WorkOrder = workOrder,
                            Task = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value),
                            AccountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value),
                            EquipmentReference = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                            StatisticType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value),
                            Value = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value),
                        };


                        if (equipmentHire.WoProjectIndicator.Equals("W") && !string.IsNullOrWhiteSpace(equipmentHire.WorkOrder))
                        {
                            if (string.IsNullOrWhiteSpace(equipmentHire.Task) && cbAutoTaskAssigment.Checked)
                                equipmentHire.Task = "001";
                            else if (!string.IsNullOrWhiteSpace(equipmentHire.Task) && equipmentHire.Task.Length >= 1 && equipmentHire.Task.Length < 3)
                                equipmentHire.Task = equipmentHire.Task.PadLeft(3, '0');
                        }


                        EquipmentHireActions.DeleteEquipmentHire(urlService, opSheet, equipmentHire);

                        _cells.GetCell(resultColumn, i).Value = "ELIMINADO";
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:DeleteEquipmentHire()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    finally
                    {
                        i++;
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:DeleteEquipmentHire()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void DeleteMse850LabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = Mse850TitleRow;
                var resultColumn = Mse850ResultColumn;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = titleRow + 1;
                //recorro la lista de empleados de forma vertical
                while (!_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2).Equals(""))
                {
                    try
                    {
                        //valores de listas
                        var transactionDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                        var employeeId = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                        var interDistrict = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                        var project = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value);
                        var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                        var woTask = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);
                        var equipmentRef = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value);
                        var equipmentNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value);
                        var laborClass = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value);
                        var earningCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value);
                        var hours = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value);
                        var overtimeIndicator = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value);
                        var accountCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value);

                        //obtengo solo el código sin la descripción
                        if (earningCode != null && earningCode.Contains(" - "))
                            earningCode = earningCode.Substring(0, earningCode.IndexOf(" - ", StringComparison.Ordinal));
                        if (laborClass != null && laborClass.Contains(" - "))
                            laborClass = laborClass.Substring(0, laborClass.IndexOf(" - ", StringComparison.Ordinal));

                        //proceso por empleado
                        var employee = new LabourEmployee
                        {
                            TransactionDate = transactionDate,
                            Employee = employeeId,
                            InterDistrictCode = interDistrict,
                            Project = project,
                            WorkOrder = workOrder,
                            WorkOrderTask = woTask,
                            EquipmentRef = equipmentRef,
                            EquipmentNo = equipmentNo,
                            LabourClass = laborClass,
                            EarnCode = earningCode,
                            LabourCostingHours = hours,
                            OvertimeInd = MyUtilities.IsTrue(overtimeIndicator),
                            AccountCode = accountCode,
                        };

                        //                        if (!string.IsNullOrWhiteSpace(employee.WorkOrder))
                        //                        {
                        //                            if (string.IsNullOrWhiteSpace(employee.WorkOrderTask) && cbAutoTaskAssigment.Checked)
                        //                                employee.WorkOrderTask = "001";
                        //                            if (!string.IsNullOrWhiteSpace(employee.WorkOrderTask) && employee.WorkOrderTask.Length >= 1 && employee.WorkOrderTask.Length < 3)
                        //                                employee.WorkOrderTask = employee.WorkOrderTask.PadLeft(3, '0');
                        //                        }

                        var reply = LabourEmployeeActions.DeleteEmployeeMse(urlService, opSheet, employee);

                        if (reply != null && reply.errors.Length == 0)
                        {
                            _cells.GetCell(resultColumn, i).Value = "ELIMINADO";
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn, i).Select();
                        }
                        else
                        {
                            string errorMessage;
                            if (reply == null)
                                errorMessage = "NO SE HA ENCONTRADO UN REGISTRO PARA ELIMINAR";
                            else if (reply != null && reply.errors.Length > 0)
                                errorMessage = string.Join(",", reply.errors.Select(p => p.messageText).ToArray());
                            else
                                errorMessage = "SE HA PRODUCIDO UN ERROR AL INTENTAR ELIMINAR UN REGISTRO";
                            _cells.GetCell(resultColumn, i).Value = errorMessage;
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                            _cells.GetCell(resultColumn, i).Select();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:DeleteMso850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(resultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn, i).Select();
                    }
                    finally
                    {
                        i++;
                    }
                } //--while de registros
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:DeleteMse850LabourCost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }


        public void LoadElecsaLabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = ElecsaTitleRow;
                var titleColumn = ElecsaTitleColumn;
                var resultColumn = ElecsaResultColumn;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, titleRow - 3).Value2);
                var loadType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, titleRow - 3).Value2);

                var i = titleRow + 1;
                const string employeeId = "CONPBVELE";
                const string earningCode = "001";
                var transDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, titleRow - 2).Value2);

                if (_cells.GetEmptyIfNull(loadType).Equals("LABOR") ||
                    _cells.GetEmptyIfNull(loadType).Equals("LABOR Y DURACIÓN"))
                {
                    //recorro la lista de tareas de forma vertical
                    while (
                        !(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2).Equals("") &&
                          _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals("")))
                    {
                        var errorFlag = false;

                        var task = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value2);
                        task = (task == null)  && cbAutoTaskAssigment.Checked ? "001": task;
                        //el número de tarea debe tener tres dígitos 001
                        if (string.IsNullOrWhiteSpace(task) && task.Length >= 1 && task.Length < 3)
                            task = task.PadLeft(3, '0');
                        var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value2);
                        var costCenter = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value2);

                        var j = titleColumn;
                        //recorro la lista de empleados de forma horizontal
                        while (!_cells.GetEmptyIfNull(_cells.GetCell(j, titleRow).Value2).Equals("RESULTADO"))
                        {
                            try
                            {
                                var hours = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(j, i).Value);
                                if (hours == null)
                                    continue;
                                var laborClass =
                                    _cells.GetNullIfTrimmedEmpty(_cells.GetCell(j, titleRow - 1).Value2);
                                if (laborClass != null && laborClass.Contains(" - "))
                                    laborClass = laborClass.Substring(0,
                                        laborClass.IndexOf(" - ", StringComparison.Ordinal));
                                //proceso por empleado
                                var employee = new LabourEmployee
                                {
                                    TransactionDate = transDate,
                                    LabourCostingHours = hours,
                                    AccountCode = costCenter,
                                    WorkOrder = workOrder,
                                    WorkOrderTask = task,
                                    Employee = employeeId,
                                    EarnCode = earningCode,
                                    LabourClass = laborClass
                                };

                                var reply = LabourEmployeeActions.CreateEmployeeMse(urlService, opSheet, employee, false);

                                var errorMessage = "";
                                if (reply?.errors != null && reply.errors.Length > 0)
                                {
                                    errorMessage = reply.errors.Aggregate(errorMessage, (current, err) => current + err.messageText);
                                    throw new Exception(errorMessage);
                                }

                                _cells.GetCell(j, i).ClearComments();
                                _cells.GetCell(j, i).Style = StyleConstants.Success;
                                _cells.GetCell(j, i).Select();
                            }
                            catch (Exception ex)
                            {
                                errorFlag = true;
                                Debugger.LogError("RibbonEllipse:LoadElecsaLabourCost()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                                _cells.GetCell(j, i).ClearComments();
                                _cells.GetCell(j, i).AddComment(ex.Message);
                                _cells.GetCell(j, i).Style = StyleConstants.Error;
                                _cells.GetCell(j, i).Select();
                            }
                            finally
                            {
                                j++;
                            }
                        } //--while de empleados
                        if (errorFlag)
                        {
                            _cells.GetCell(resultColumn, i).Value = "Se han encontrado algunos errores";
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Error;
                            _cells.GetCell(resultColumn, i).Select();
                        }
                        else
                        {
                            _cells.GetCell(resultColumn, i).Value = "OK";
                            _cells.GetCell(resultColumn, i).Style = StyleConstants.Success;
                            _cells.GetCell(resultColumn, i).Select();
                        }
                        i++;
                    } //--while de tareas
                }

                //Si hay cargue de duración y/o comentarios
                if (!_cells.GetEmptyIfNull(loadType).Equals("DURACIÓN") &&
                    !_cells.GetEmptyIfNull(loadType).Equals("LABOR Y DURACIÓN")) return;

                i = titleRow + 1;
                var opWo = new EllipseWorkOrdersClassLibrary.WorkOrderService.OperationContext
                {
                    district = _frmAuth.EllipseDstrct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                //recorro la lista de tareas de forma vertical para duración
                while (!(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2).Equals("") && _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals("")))
                {
                    try
                    {
                        var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value2);

                        var durationCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(resultColumn + 1, i).Value2);
                        var startHour = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(resultColumn + 2, i).Value2);
                        var endHour = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(resultColumn + 3, i).Value2);
                        var wo = WorkOrderActions.GetNewWorkOrderDto(long.TryParse("" + workOrder, out long number1) ? ("" + workOrder).PadLeft(8, '0') : workOrder);
                        var completeCommentToAppend = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(resultColumn + 4, i).Value2);

                        if (_cells.GetNullIfTrimmedEmpty(durationCode) != null)
                        {
                            var duration = new WorkOrderDuration
                            {
                                jobDurationsDate = transDate,
                                jobDurationsCode = durationCode,
                                jobDurationsStart = startHour,
                                jobDurationsFinish = endHour
                            };

                            WorkOrderActions.CreateWorkOrderDuration(urlService, opWo, districtCode, wo, duration);
                        }

                        if (_cells.GetNullIfTrimmedEmpty(completeCommentToAppend) != null)
                        {
                            var stdTextId = "CW" + districtCode + wo.prefix + wo.no;

                            var stdTextCopc = StdText.GetCustomOpContext(opWo.district, opWo.position, opWo.maxInstances, opWo.returnWarnings);
                            var woCompleteComment = StdText.GetText(urlService, stdTextCopc, stdTextId);

                            StdText.SetText(urlService, stdTextCopc, stdTextId, woCompleteComment + "\n" + completeCommentToAppend);
                        }


                        _cells.GetCell(resultColumn + 5, i).Value = "OK";
                        _cells.GetCell(resultColumn + 5, i).Style = StyleConstants.Success;
                        _cells.GetCell(resultColumn + 5, i).Select();

                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(resultColumn + 5, i).Value = "ERROR" + ex.Message;
                        _cells.GetCell(resultColumn + 5, i).Style = StyleConstants.Error;
                        _cells.GetCell(resultColumn + 5, i).Select();
                    }
                    finally
                    {
                        i++;
                    }
                } //--while de tareas
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void CleanLabourTable()
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mso850))//Deprecated
                _cells.ClearTableRange(TableNameMso850);//Deprecated
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mse850))
                _cells.ClearTableRange(TableNameMse850);
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Elecsa))
                _cells.ClearTableRange(TableNameElecsa);
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Default))
                _cells.ClearTableRange(TableNameDefault);
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.EquipmentHire))
                _cells.ClearTableRange(TableNameEquipmentHire);

        }

        

        /// <summary>
        /// Establece el resultado de búsqueda de la descripción de una orden después de que esta es escrita
        /// </summary>
        /// <param name="target"></param>
        void GetDefaultWorkOrderDescriptionChangedValue(Excel.Range target)
        {
            try
            {
                if (_cells.GetNullIfTrimmedEmpty(target.Text) == null)
                {
                    _cells.GetCell(target.Column, target.Row + 1).Value = "";
                    return;
                }
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.GetCell(target.Column, target.Row + 1).Value = "Buscando Orden...";
                var district = "" + _cells.GetCell("B3").Value;

                string woNo = _cells.GetEmptyIfNull(target.Value2);
                if (woNo.All(char.IsDigit) && woNo.Length < 8)
                    woNo = woNo.PadLeft(8, '0');

                var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, "" + district, woNo);

                if (wo != null)
                    _cells.GetCell(target.Column, target.Row + 1).Value = "" + wo.workOrderDesc;
                else
                    _cells.GetCell(target.Column, target.Row + 1).Value = "Orden no encontrada";

            }
            catch (NullReferenceException)
            {
                _cells.GetCell(target.Column, target.Row + 1).Value = "No fue Posible Obtener Informacion!";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        /// <summary>
        /// Establece el resultado de búsqueda de la descripción de un equipo después de que este es escrita
        /// </summary>
        /// <param name="target"></param>
        /// <param name="changedRanges"></param>
        void GetTableChangedValue(Excel.Range target, ListRanges changedRanges)//Excel.Range target)
        {
            switch (target.Column)
            {
                case 2:
                    try
                    {
                        if (_cells.GetNullIfTrimmedEmpty(target.Text) == null)
                        {
                            _cells.GetCell(target.Column + 2, target.Row).Value = "";
                            return;
                        }
                        _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                        _cells.GetCell(target.Column + 2, target.Row).Value = "Buscando Orden...";
                        var district = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, ElecsaTitleRow - 3).Value2);

                        string woNo = _cells.GetEmptyIfNull(target.Value2);
                        if (woNo.All(char.IsDigit) && woNo.Length < 8)
                            woNo = woNo.PadLeft(8, '0');

                        var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, "" + district, woNo);

                        if (wo != null)
                            _cells.GetCell(target.Column + 2, target.Row).Value = "" + wo.workOrderDesc;
                        else
                            _cells.GetCell(target.Column + 2, target.Row).Value = "Orden no encontrada";

                    }
                    catch (NullReferenceException)
                    {
                        _cells.GetCell(target.Column + 2, target.Row).Value = "No fue Posible Obtener Informacion!";
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.Message);
                    }
                    break;
            }
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

        private void btnReviewWorkOrder_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Default))
                {
                    var titleRow = DefaultTitleRow;

                    _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                    var district = "" + _cells.GetCell("B3").Value;
                    for (var i = 0; i < OtFields; i++)
                    {
                        var target = _cells.GetCell(DefaultStartCol + i * NroEle, titleRow - 2);
                        var woNo = _cells.GetNullIfTrimmedEmpty(target.Value2);
                        _cells.GetCell(target.Column, target.Row + 1).Value = "";

                        if (woNo == null) continue;
                        _cells.GetCell(target.Column, target.Row + 1).Value = "Buscando Orden...";
                        var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, "" + district, woNo);
                        if (wo != null)
                            _cells.GetCell(target.Column, target.Row + 1).Value = "" + wo.workOrderDesc;
                        else
                            _cells.GetCell(target.Column, target.Row + 1).Value = "Orden no encontrada";
                    }
                }
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Elecsa))
                {
                    var titleRow = ElecsaTitleRow;

                    _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                    var district = "" + _cells.GetCell("B3").Value;
                    var i = titleRow + 1;
                    while (_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value2) != null)
                    {
                        var target = _cells.GetCell(2, i);
                        var woNo = _cells.GetNullIfTrimmedEmpty(target.Value2);
                        _cells.GetCell(target.Column + 2, target.Row).Value = "";
                        i++;

                        if (woNo == null) continue;
                        _cells.GetCell(target.Column + 2, target.Row).Value = "Buscando Orden...";
                        var wo = WorkOrderActions.FetchWorkOrder(_eFunctions, "" + district, woNo);
                        if (wo != null)
                            _cells.GetCell(target.Column + 2, target.Row).Value = "" + wo.workOrderDesc;
                        else
                            _cells.GetCell(target.Column + 2, target.Row).Value = "Orden no encontrada";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnDeleteLaborSheet_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                if (!_cells.IsDecimalDotSeparator())
                    if (MessageBox.Show(@"El separador de decimales configurado actualmente no es el punto. Usar un separador de decimales diferente puede generar errores al momento de cargar valores numéricos. ¿Está seguro que desea continuar?", @"ALERTA DE SEPARADOR DE DECIMALES", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mse850))
                    _thread = new Thread(DeleteMse850LabourCost);
                else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.EquipmentHire))
                    _thread = new Thread(DeleteEquipmentHire);
                else
                {
                    MessageBox.Show(@"La hoja de Excel no tiene el formato válido para el cargue de labor");
                    return;
                }
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:DeleteLaborSheet()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        
    }
    
    
}
