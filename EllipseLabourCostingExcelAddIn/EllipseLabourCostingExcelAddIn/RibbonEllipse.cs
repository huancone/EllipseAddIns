using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using System.Web.Services.Ellipse;
using EllipseWorkOrdersClassLibrary;
using EllipseStdTextClassLibrary;
using System.Threading;
using EllipseLabourCostingExcelAddIn.LabourCostingTransService;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using OperationContext = EllipseLabourCostingExcelAddIn.LabourCostingTransService.OperationContext;
// ReSharper disable LoopCanBeConvertedToQuery
namespace EllipseLabourCostingExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;
        const string SheetName01 = "Labour";
        private const string ValidationSheetName = "ValidationSheetLabour";

        const int DefaultStartCol = 3;
        const int NroEle = 3;//cantidad de elementos por orden a ingresar (HRS, TSK, Earning Code)
        const int DefaultTitleRow = 9;
        const int Mso850TitleRow = 5;//Deprecated
        const int Mse850TitleRow = 5;
        const int ElecsaTitleRow = 6;
        const int ElecsaTitleColumn = 6;
        private const int Mso850ResultColumn = 18;//Deprecated
        private const int Mse850ResultColumn = 14;
        private const int ElecsaResultColumn = 30;
        private const string TableNameMso850 = "Mso850Table";//Deprecated
        private const string TableNameMse850 = "Mse850Table";
        private const string TableNameDefault = "LabourDefaultTable";
        private const string TableNameElecsa = "ElecsaTable";

        private const int OtFields = 20;

        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }
        }

        private void btnFormatHeader_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetHeaderData();
        }

        private void btnFormatDefault_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Contains(SheetName01))
            {
                var groupName = "" + _cells.GetCell("B4").Value;

                if (!groupName.Equals(""))
                    FormatGroupDefault(groupName);
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
        public void FormatGroupDefault(string groupName)
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
                var sqlQuery = Queries.GetGroupEmployeesQuery(groupName, _eFunctions.dbReference, _eFunctions.dbLink);
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
                    _cells.GetCell(DefaultStartCol + i * NroEle, DefaultTitleRow - 3).Value = "OT " + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle, DefaultTitleRow - 3).Style =
                        _cells.GetStyle(StyleConstants.TitleOptional);
                    _cells.MergeCells(DefaultStartCol + i * NroEle, DefaultTitleRow - 3,
                        DefaultStartCol + i * NroEle + (NroEle - 1), DefaultTitleRow - 3);
                    //Número OT
                    _cells.GetCell(DefaultStartCol + i * NroEle, DefaultTitleRow - 2).Style =
                        _cells.GetStyle(StyleConstants.Select);
                    vstoSheet.Controls.Remove("SeekOrder" + i);
                    var orderNameRange =
                        vstoSheet.Controls.AddNamedRange(
                            _cells.GetCell(DefaultStartCol + i * NroEle, DefaultTitleRow - 2),
                            "SeekOrder" + i);
                    orderNameRange.Change += GetDefaultWorkOrderDescriptionChangedValue;
                    _cells.GetRange(DefaultStartCol + i * NroEle, DefaultTitleRow - 2,
                        DefaultStartCol + i * NroEle + (NroEle - 1), DefaultTitleRow - 2)
                        .NumberFormat = NumberFormatConstants.Text;
                    _cells.MergeCells(DefaultStartCol + i * NroEle, DefaultTitleRow - 2,
                        DefaultStartCol + i * NroEle + (NroEle - 1), DefaultTitleRow - 2);
                    //Descripción OT
                    _cells.GetCell(DefaultStartCol + i * NroEle, DefaultTitleRow - 1).Style =
                        _cells.GetStyle(StyleConstants.Option);
                    _cells.MergeCells(DefaultStartCol + i * NroEle, DefaultTitleRow - 1,
                        DefaultStartCol + i * NroEle + (NroEle - 1), DefaultTitleRow - 1);
                    //Cada componente
                    _cells.GetCell(DefaultStartCol + i * NroEle + 0, DefaultTitleRow).Value = "HRS_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 0, DefaultTitleRow).AddComment("hh.mm");
                    _cells.GetCell(DefaultStartCol + i * NroEle + 0, DefaultTitleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleRequired);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 1, DefaultTitleRow).Value = "TASK_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 1, DefaultTitleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleRequired);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 2, DefaultTitleRow).Value = "ECODE_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i * NroEle + 2, DefaultTitleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleOptional);

                }
                _cells.GetRange(DefaultStartCol, DefaultTitleRow + 1, DefaultStartCol + OtFields * NroEle,
                    DefaultTitleRow + 1)
                    .NumberFormat = NumberFormatConstants.Text;
                //Validación para los Earning Codes
                var itemList = _eFunctions.GetItemCodes("EA");
                var validationList = itemList.Select(item => item.code + " - " + item.description).ToList();

                //creo la validación y la asigno al primer elemento
                _cells.SetValidationList(_cells.GetCell(DefaultStartCol + 2, DefaultTitleRow + 1), validationList,
                    ValidationSheetName, 1);
                var validation = _cells.GetCell(DefaultStartCol + 2, DefaultTitleRow + 1).Validation;
                //asigno la validación al resto de los elementos
                for (var j = 1; j < OtFields; j++)
                {
                    _cells.GetCell(DefaultStartCol + j * NroEle + 2, DefaultTitleRow + 1).Validation.Delete();
                    _cells.GetCell(DefaultStartCol + j * NroEle + 2, DefaultTitleRow + 1)
                        .Validation.Add((Excel.XlDVType)validation.Type, validation.AlertStyle, validation.Operator,
                            validation.Formula1, validation.Formula2);
                }
                _cells.FormatAsTable(
                    _cells.GetRange(1, DefaultTitleRow, DefaultStartCol + OtFields * NroEle - 1, DefaultTitleRow + 1),
                    TableNameDefault);
                var lengthEmployees = 0;
                if (drEmployees != null && !drEmployees.IsClosed && drEmployees.HasRows)
                {
                    while (drEmployees.Read())
                    {
                        _cells.GetCell(1, DefaultTitleRow + 1 + lengthEmployees).Value =
                            drEmployees["NOMBRE"].ToString().Trim();
                        _cells.GetCell(2, DefaultTitleRow + 1 + lengthEmployees).Value =
                            drEmployees["CEDULA"].ToString().Trim();
                        lengthEmployees++;
                    }
                }
                else
                    MessageBox.Show(@"No se han encontrado datos para el modelo especificado");

                _eFunctions.CloseConnection();
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

                _cells.GetRange(1, Mse850TitleRow, Mse850ResultColumn - 1, Mse850TitleRow).Style = StyleConstants.TitleRequired;
                _cells.GetRange(1, Mse850TitleRow + 1, Mse850ResultColumn - 1, Mse850TitleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(1, Mse850TitleRow).Value = "Trans.Date";
                _cells.GetCell(1, Mse850TitleRow).AddComment("yyyyMMss");
                _cells.GetCell(2, Mse850TitleRow).Value = "EmployeeId";
                _cells.GetCell(3, Mse850TitleRow).Value = "InterDistrict";
                _cells.GetCell(3, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, Mse850TitleRow).Value = "Project";
                _cells.GetCell(4, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, Mse850TitleRow).Value = "WorkOrder";
                _cells.GetCell(6, Mse850TitleRow).Value = "WOTask";
                _cells.GetCell(6, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, Mse850TitleRow).Value = "EquipmentRef";
                _cells.GetCell(8, Mse850TitleRow).Value = "EquipmentNo";
                _cells.GetCell(8, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(9, Mse850TitleRow).Value = "LaborClass";
                //_cells.SetValidationList(_cells.GetCell(9, Mse850TitleRow + 1), laborClassList, ValidationSheetName, 2);
                _cells.GetCell(9, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, Mse850TitleRow).Value = "EarningCode";
                //_cells.SetValidationList(_cells.GetCell(10, Mse850TitleRow + 1), earningCodeList, ValidationSheetName, 3);
                _cells.GetCell(10, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, Mse850TitleRow).Value = "Hours";
                _cells.GetCell(11, Mse850TitleRow).AddComment("hh.mm");
                _cells.GetCell(12, Mse850TitleRow).Value = "OvertimeIndicator";
                _cells.GetCell(12, Mse850TitleRow).AddComment("Y/N");
                _cells.GetCell(12, Mse850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(13, Mse850TitleRow).Value = "AccountCode";
                _cells.GetCell(13, Mse850TitleRow).Style = StyleConstants.TitleOptional;

                _cells.GetCell(Mse850ResultColumn, Mse850TitleRow).Value = "RESULTADO";
                _cells.GetCell(Mse850ResultColumn, Mse850TitleRow).Style = StyleConstants.TitleResult;
                _cells.FormatAsTable(_cells.GetRange(1, Mse850TitleRow, Mse850ResultColumn, Mse850TitleRow + 1), TableNameMse850);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
                var compCodeList = itemList1.Select(item => item.code + " - " + item.description).ToList();

                var itemList2 = _eFunctions.GetItemCodes("EA");//earning codes
                var earningCodeList = itemList2.Select(item => item.code + " - " + item.description).ToList();

                var itemList3 = _eFunctions.GetItemCodes("LC");//laborclass codes
                var laborClassList = itemList3.Select(item => item.code + " - " + item.description).ToList();

                _cells.GetRange(1, Mso850TitleRow, Mso850ResultColumn - 1, Mso850TitleRow).Style = StyleConstants.TitleRequired;
                _cells.GetRange(1, Mso850TitleRow + 1, Mso850ResultColumn - 1, Mso850TitleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(1, Mso850TitleRow).Value = "TRAN_DATE";
                _cells.GetCell(1, Mso850TitleRow).AddComment("yyyyMMss");
                _cells.GetCell(2, Mso850TitleRow).Value = "ORD_HRS";
                _cells.GetCell(2, Mso850TitleRow).AddComment("hh.mm");
                _cells.GetCell(3, Mso850TitleRow).Value = "OT_HRS";
                _cells.GetCell(3, Mso850TitleRow).AddComment("hh.mm");
                _cells.GetCell(3, Mso850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, Mso850TitleRow).Value = "VALUE";
                _cells.GetCell(4, Mso850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, Mso850TitleRow + 1).Style = StyleConstants.Disabled;
                _cells.GetCell(5, Mso850TitleRow).Value = "INT_DSTRCT_CDE";
                _cells.GetCell(6, Mso850TitleRow).Value = "ACCOUNT";
                _cells.GetCell(7, Mso850TitleRow).Value = "STATUS";
                _cells.GetCell(7, Mso850TitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, Mso850TitleRow + 1).Style = StyleConstants.Disabled;
                _cells.GetCell(8, Mso850TitleRow).Value = "WOP_IND";
                _cells.GetCell(8, Mso850TitleRow).AddComment("P: Proyecto \nW: WorkOrder");
                _cells.SetValidationList(_cells.GetCell(8, Mso850TitleRow + 1), new List<string> { "P", "W" }, ValidationSheetName, 1);
                _cells.GetCell(9, Mso850TitleRow).Value = "WO_PROJ";
                _cells.GetCell(10, Mso850TitleRow).Value = "TASK";
                _cells.GetCell(11, Mso850TitleRow).Value = "EMP_ID";
                _cells.GetCell(12, Mso850TitleRow).Value = "EQUIPMENT";
                _cells.GetCell(13, Mso850TitleRow).Value = "UNITS_COMP";
                _cells.GetCell(14, Mso850TitleRow).Value = "PC_COMP";
                _cells.GetCell(15, Mso850TitleRow).Value = "CODE_COMP";
                _cells.SetValidationList(_cells.GetCell(15, Mso850TitleRow + 1), compCodeList, ValidationSheetName, 2);
                _cells.GetCell(16, Mso850TitleRow).Value = "EARN_CLASS";
                _cells.SetValidationList(_cells.GetCell(16, Mso850TitleRow + 1), earningCodeList, ValidationSheetName, 3);
                _cells.GetCell(17, Mso850TitleRow).Value = "LABOUR_CLASS";
                _cells.SetValidationList(_cells.GetCell(17, Mso850TitleRow + 1), laborClassList, ValidationSheetName, 4);


                _cells.GetCell(Mso850ResultColumn, Mso850TitleRow).Value = "RESULTADO";
                _cells.GetCell(Mso850ResultColumn, Mso850TitleRow).Style = StyleConstants.TitleResult;
                _cells.FormatAsTable(
                    _cells.GetRange(1, Mso850TitleRow, Mso850ResultColumn, Mso850TitleRow + 1),
                    TableNameMso850);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "ELECSA SÁBANA DE LABOR - ELLIPSE 8";
                _cells.GetCell("B1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells(2, 1, ElecsaResultColumn - 1, 2);

                _cells.GetCell(ElecsaResultColumn, 1).Value = "OBLIGATORIO";
                _cells.GetCell(ElecsaResultColumn, 1).Style = StyleConstants.TitleRequired;
                _cells.GetCell(ElecsaResultColumn, 2).Value = "OPCIONAL";
                _cells.GetCell(ElecsaResultColumn, 2).Style = StyleConstants.TitleOptional;
                _cells.GetCell(ElecsaResultColumn, 3).Value = "INFORMATIVO";
                _cells.GetCell(ElecsaResultColumn, 3).Style = StyleConstants.TitleInformation;

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

                _cells.GetCell(1, ElecsaTitleRow - 3).Value = "DISTRITO";
                _cells.GetCell(1, ElecsaTitleRow - 3).Style = StyleConstants.Option;
                _cells.GetCell(2, ElecsaTitleRow - 3).Value = "ICOR";
                _cells.GetCell(2, ElecsaTitleRow - 3).Style = StyleConstants.Select;
                _cells.GetCell(1, ElecsaTitleRow - 2).Value = "FECHA";
                _cells.GetCell(1, ElecsaTitleRow - 2).Style = StyleConstants.Option;
                _cells.GetCell(1, ElecsaTitleRow - 2).AddComment("YYYYMMDD");
                _cells.GetCell(2, ElecsaTitleRow - 2).Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell(2, ElecsaTitleRow - 2).Style = StyleConstants.Select;

                _cells.GetCell(3, ElecsaTitleRow - 3).Value = "TIPO CARGA";
                _cells.GetCell(3, ElecsaTitleRow - 3).Style = StyleConstants.Option;
                _cells.GetCell(4, ElecsaTitleRow - 3).Value = "LABOR Y DURACIÓN";
                _cells.GetCell(4, ElecsaTitleRow - 3).Style = StyleConstants.Select;
                var loadTypeList = new List<string>
                {
                    "LABOR",
                    "DURACIÓN",
                    "LABOR Y DURACIÓN"
                };

                _cells.GetCell(ElecsaTitleColumn - 1, ElecsaTitleRow - 2).Value = "EMPLEADOS";
                _cells.GetCell(ElecsaTitleColumn - 1, ElecsaTitleRow - 2).Style = StyleConstants.Option;
                _cells.GetCell(1, ElecsaTitleRow).Value = "CENTRO";
                _cells.GetCell(1, ElecsaTitleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(2, ElecsaTitleRow).Value = "WORKORDER";
                _cells.GetCell(2, ElecsaTitleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(3, ElecsaTitleRow).Value = "TASK";
                _cells.GetCell(4, ElecsaTitleRow).Value = "DESCRIPCIÓN";
                _cells.GetRange(2, ElecsaTitleRow, 4, ElecsaTitleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(ElecsaTitleColumn - 1, ElecsaTitleRow - 1).Value = "LABOR";
                _cells.GetCell(ElecsaTitleColumn - 1, ElecsaTitleRow - 1).Style = StyleConstants.Option;
                _cells.GetCell(ElecsaTitleColumn - 1, ElecsaTitleRow).Value = "Nro.";
                _cells.GetCell(ElecsaTitleColumn - 1, ElecsaTitleRow + 1).Style = StyleConstants.Disabled;
                for (var i = ElecsaTitleColumn; i < ElecsaResultColumn; i++)
                {
                    _cells.GetCell(i, ElecsaTitleRow).Value = "" + (i - ElecsaTitleColumn + 1);
                    _cells.GetCell(i, ElecsaTitleRow).AddComment("hh.mm");
                }
                _cells.GetRange(ElecsaTitleColumn, ElecsaTitleRow, ElecsaResultColumn - 1, ElecsaTitleRow).Style =
                    StyleConstants.TitleOptional;


                _cells.GetRange(ElecsaTitleColumn, ElecsaTitleRow - 2, ElecsaResultColumn - 1, ElecsaTitleRow - 1).Style = StyleConstants.Select;
                //validaciones de campo
                _cells.SetValidationList(_cells.GetCell(2, ElecsaTitleRow - 3), Districts.GetDistrictList(), ValidationSheetName, 1);
                _cells.SetValidationList(_cells.GetCell(4, ElecsaTitleRow - 3), loadTypeList, ValidationSheetName, 2);

                _cells.SetValidationList(_cells.GetRange(ElecsaTitleColumn, ElecsaTitleRow - 1, ElecsaResultColumn - 1, ElecsaTitleRow - 1), laborClassList, ValidationSheetName, 3);
                _cells.GetRange(ElecsaTitleColumn - 1, ElecsaTitleRow - 2, ElecsaResultColumn - 1, ElecsaTitleRow - 1).ColumnWidth = 3.57;
                _cells.GetRange(ElecsaTitleColumn - 1, ElecsaTitleRow - 2, ElecsaResultColumn - 1, ElecsaTitleRow - 1).Orientation = 90;

                _cells.GetCell(ElecsaResultColumn, ElecsaTitleRow).Value = "RESULTADO";
                _cells.GetCell(ElecsaResultColumn, ElecsaTitleRow).Style = StyleConstants.TitleResult;

                _cells.GetCell(ElecsaResultColumn + 1, ElecsaTitleRow).Value = "DURATION CODE";
                _cells.GetCell(ElecsaResultColumn + 2, ElecsaTitleRow).Value = "HR INICIAL";
                _cells.GetCell(ElecsaResultColumn + 2, ElecsaTitleRow).AddComment("hhmmss");
                _cells.GetCell(ElecsaResultColumn + 2, ElecsaTitleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(ElecsaResultColumn + 3, ElecsaTitleRow).Value = "HR FINAL";
                _cells.GetCell(ElecsaResultColumn + 3, ElecsaTitleRow).AddComment("hhmmss");
                _cells.GetCell(ElecsaResultColumn + 3, ElecsaTitleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(ElecsaResultColumn + 4, ElecsaTitleRow).Value = "COMENTARIO";
                _cells.GetCell(ElecsaResultColumn + 4, ElecsaTitleRow).AddComment("No modifica el comentario de cierre, solo adiciona al comentario existente");
                _cells.GetCell(ElecsaResultColumn + 4, ElecsaTitleRow + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetRange(ElecsaResultColumn + 1, ElecsaTitleRow, ElecsaResultColumn + 4, ElecsaTitleRow).Style = StyleConstants.TitleRequired;


                _cells.GetCell(ElecsaResultColumn + 5, ElecsaTitleRow).Value = "RESULTADO DUR.";
                _cells.GetCell(ElecsaResultColumn + 5, ElecsaTitleRow).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, ElecsaTitleRow + 1, ElecsaResultColumn + 5, ElecsaTitleRow + 1).NumberFormat =
                    NumberFormatConstants.Text;

                //búsquedas especiales de tabla
                var table = _cells.FormatAsTable(_cells.GetRange(1, ElecsaTitleRow, ElecsaResultColumn + 5, ElecsaTitleRow + 1), TableNameElecsa);
                var tableObject = Globals.Factory.GetVstoObject(table);
                tableObject.Change += GetTableChangedValue;

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _cells.SetWorksheetVisibility(ValidationSheetName, false);

            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
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

                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                //Deprecated
                //var opSheet = new Screen.OperationContext
                //{
                //    district = _frmAuth.EllipseDsct,
                //    position = _frmAuth.EllipsePost,
                //    maxInstances = 100,
                //    returnWarnings = Debugger.DebugWarnings
                //};
                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
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

                    while ("" + _cells.GetCell(j, DefaultTitleRow - 2).Value != "")
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

                                employee.WorkOrder = _cells.GetEmptyIfNull(_cells.GetCell(j, DefaultTitleRow - 2).Value2);

                                LoadEmployeeMse(urlService, opSheet, employee, cbReplaceExisting.Checked);

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

                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = Mse850TitleRow + 1;
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

                        var reply = LoadEmployeeMse(urlService, opSheet, employee, cbReplaceExisting.Checked);

                        if (reply.errors.Length == 0)
                        {
                            _cells.GetCell(Mse850ResultColumn, i).Value = "SUCCESS";
                            _cells.GetCell(Mse850ResultColumn, i).Style = StyleConstants.Success;
                            _cells.GetCell(Mse850ResultColumn, i).Select();
                        }
                        else
                        {
                            _cells.GetCell(Mse850ResultColumn, i).Value = string.Join(",", reply.errors.Select(p => p.messageText).ToArray());
                            _cells.GetCell(Mse850ResultColumn, i).Style = StyleConstants.Error;
                            _cells.GetCell(Mse850ResultColumn, i).Select();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(Mse850ResultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(Mse850ResultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(Mse850ResultColumn, i).Select();
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

                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = Mso850TitleRow + 1;
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


                        LoadEmployeeMso(opSheet, employee, cbReplaceExisting.Checked);

                        _cells.GetCell(Mso850ResultColumn, i).Value = "SUCCESS";
                        _cells.GetCell(Mso850ResultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(Mso850ResultColumn, i).Select();
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(Mso850ResultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(Mso850ResultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(Mso850ResultColumn, i).Select();
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

        public void DeleteMse850LabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = Mse850TitleRow + 1;
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

                        var reply = DeleteEmployeeMse(urlService, opSheet, employee);

                        if (reply != null && reply.errors.Length == 0)
                        {
                            _cells.GetCell(Mse850ResultColumn, i).Value = "ELIMINADO";
                            _cells.GetCell(Mse850ResultColumn, i).Style = StyleConstants.Success;
                            _cells.GetCell(Mse850ResultColumn, i).Select();
                        }
                        else
                        {
                            _cells.GetCell(Mse850ResultColumn, i).Value = string.Join(",", reply.errors.Select(p => p.messageText).ToArray());
                            _cells.GetCell(Mse850ResultColumn, i).Style = StyleConstants.Error;
                            _cells.GetCell(Mse850ResultColumn, i).Select();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                        _cells.GetCell(Mse850ResultColumn, i).Value = "ERROR: " + ex.Message;
                        _cells.GetCell(Mse850ResultColumn, i).Style = StyleConstants.Error;
                        _cells.GetCell(Mse850ResultColumn, i).Select();
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


        public void LoadElecsaLabourCost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                //Deprecated
                //var opSheet = new Screen.OperationContext
                //{
                //    district = _frmAuth.EllipseDsct,
                //    position = _frmAuth.EllipsePost,
                //    maxInstances = 100,
                //    returnWarnings = Debugger.DebugWarnings
                //};

                var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = Debugger.DebugWarnings
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, ElecsaTitleRow - 3).Value2);
                var loadType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, ElecsaTitleRow - 3).Value2);

                var i = ElecsaTitleRow + 1;
                const string employeeId = "CONPBVELE";
                const string earningCode = "001";
                var transDate = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, ElecsaTitleRow - 2).Value2);

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
                        task = task && cbAutoTaskAssigment.Checked ?? "001";
                        //el número de tarea debe tener tres dígitos 001
                        if (string.IsNullOrWhiteSpace(task) && task.Length >= 1 && task.Length < 3)
                            task = task.PadLeft(3, '0');
                        var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value2);
                        var costCenter = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value2);

                        var j = ElecsaTitleColumn;
                        //recorro la lista de empleados de forma horizontal
                        while (!_cells.GetEmptyIfNull(_cells.GetCell(j, ElecsaTitleRow).Value2).Equals("RESULTADO"))
                        {
                            try
                            {
                                var hours = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(j, i).Value);
                                if (hours == null)
                                    continue;
                                var laborClass =
                                    _cells.GetNullIfTrimmedEmpty(_cells.GetCell(j, ElecsaTitleRow - 1).Value2);
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

                                var reply = LoadEmployeeMse(urlService, opSheet, employee, false);

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
                            _cells.GetCell(ElecsaResultColumn, i).Value = "Se han encontrado algunos errores";
                            _cells.GetCell(ElecsaResultColumn, i).Style = StyleConstants.Error;
                            _cells.GetCell(ElecsaResultColumn, i).Select();
                        }
                        else
                        {
                            _cells.GetCell(ElecsaResultColumn, i).Value = "OK";
                            _cells.GetCell(ElecsaResultColumn, i).Style = StyleConstants.Success;
                            _cells.GetCell(ElecsaResultColumn, i).Select();
                        }
                        i++;
                    } //--while de tareas
                }

                //Si hay cargue de duración y/o comentarios
                if (!_cells.GetEmptyIfNull(loadType).Equals("DURACIÓN") &&
                    !_cells.GetEmptyIfNull(loadType).Equals("LABOR Y DURACIÓN")) return;

                i = ElecsaTitleRow + 1;
                var opWo = new EllipseWorkOrdersClassLibrary.WorkOrderService.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
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
                        long number1;

                        var durationCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 1, i).Value2);
                        var startHour = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 2, i).Value2);
                        var endHour = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 3, i).Value2);
                        var wo = WorkOrderActions.GetNewWorkOrderDto(long.TryParse("" + workOrder, out number1) ? ("" + workOrder).PadLeft(8, '0') : workOrder);
                        var completeCommentToAppend = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 4, i).Value2);

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


                        _cells.GetCell(ElecsaResultColumn + 5, i).Value = "OK";
                        _cells.GetCell(ElecsaResultColumn + 5, i).Style = StyleConstants.Success;
                        _cells.GetCell(ElecsaResultColumn + 5, i).Select();

                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ElecsaResultColumn + 5, i).Value = "ERROR" + ex.Message;
                        _cells.GetCell(ElecsaResultColumn + 5, i).Style = StyleConstants.Error;
                        _cells.GetCell(ElecsaResultColumn + 5, i).Select();
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

        }

        public LabourCostingTransServiceResult LoadEmployeeMse(string urlService, OperationContext opContext, LabourEmployee labourEmployee, bool replaceExisting = true)
        {
            var proxyLt = new LabourCostingTransService.LabourCostingTransService { Url = urlService + "/LabourCostingTrans" };

            var requestLt = new LabourCostingTransDTO
            {
                transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)), int.Parse(labourEmployee.TransactionDate.Substring(4, 2)), int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                transactionDateSpecified = true,
                labourCostingHours = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingHours)? Convert.ToDecimal(labourEmployee.LabourCostingHours): default(decimal),
                labourCostingHoursSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingHours),
                labourCostingValue = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingValue)? Convert.ToDecimal(labourEmployee.LabourCostingValue): default(decimal),
                labourCostingValueSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingValue),
                interDistrictCode = labourEmployee.InterDistrictCode,
                accountCode = labourEmployee.AccountCode,
                postingStatus = labourEmployee.PostingStatus,
                project = labourEmployee.Project,
                workOrder = labourEmployee.WorkOrder,
                workOrderTask = labourEmployee.WorkOrderTask,
                employee = labourEmployee.Employee,
                equipmentNo = labourEmployee.EquipmentNo,
                equipmentReference = labourEmployee.EquipmentRef,
                percentComplete = !string.IsNullOrWhiteSpace(labourEmployee.PercentComplete)? Convert.ToDecimal(labourEmployee.PercentComplete): default(decimal),
                percentCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.PercentComplete),
                earnCode = labourEmployee.EarnCode,
                labourClass = labourEmployee.LabourClass,
                overtimeInd = labourEmployee.OvertimeInd,
                overtimeIndSpecified = true,
                unitsComplete = !string.IsNullOrWhiteSpace(labourEmployee.UnitsComplete)? Convert.ToDecimal(labourEmployee.UnitsComplete): default(decimal),
                unitsCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.UnitsComplete),
                completedCode = labourEmployee.CompletedCode
            };

            //Search Existing
            if (replaceExisting)
            {
                var requestSearch = new LabourCostingTransSearchParam
                {
                    transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)),
                        int.Parse(labourEmployee.TransactionDate.Substring(4, 2)),
                        int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                    transactionDateSpecified = true,
                    employee = labourEmployee.Employee,
                    project = labourEmployee.Project,
                    workOrder = labourEmployee.WorkOrder,
                    workOrderTask = labourEmployee.WorkOrderTask
                };
                var searchRestartDto = new LabourCostingTransDTO();
                //La búsqueda solo toma en cuenta la fecha de transacción y el employee id
                var replySearch = proxyLt.search(opContext, requestSearch, searchRestartDto);
                //Existe un elemento
                if (replySearch == null || replySearch.Length < 1) return proxyLt.create(opContext, requestLt);
                foreach (var replyItem in replySearch)
                {
                    //Las comparaciones deben hacerse con LPAD para poder establecer bien las comparaciones númericas que trae ellipse en su información
                    var equalTranDate =
                        replyItem.labourCostingTransDTO.transactionDate.Equals(requestSearch.transactionDate);
                    var equalEmployee = replyItem.labourCostingTransDTO.employee.PadLeft(20, '0')
                        .Equals(requestSearch.employee.PadLeft(20, '0'));
                    //posibles nulos de WorkOrders y/o Projects
                    var equalWo = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrder) && replyItem
                                      .labourCostingTransDTO.workOrder.PadLeft(20, '0')
                                      .Equals(requestSearch.workOrder.PadLeft(20, '0'));

                    string itemTaskNo = string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrderTask) ? "000" : replyItem.labourCostingTransDTO.workOrderTask;
                    var equalTask = itemTaskNo.Equals(!string.IsNullOrWhiteSpace(requestSearch.workOrderTask) ? requestSearch.workOrderTask.PadLeft(3, '0') : "000");
                    var equalProject = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.project) &&
                                       replyItem.labourCostingTransDTO.project.PadLeft(20, '0')
                                           .Equals(requestSearch.project.PadLeft(20, '0'));

                    if (!equalTranDate || !equalEmployee || ((!equalWo || !equalTask) && !equalProject)) continue;
                    var result = proxyLt.delete(opContext, replySearch[0].labourCostingTransDTO);
                    break;
                }

                //se envía la acción
                //return proxyLt.multipleCreate(opContext, multipleRequestLt);
                return proxyLt.create(opContext, requestLt);
            }

            return proxyLt.create(opContext, requestLt);


        }

        public LabourCostingTransServiceResult DeleteEmployeeMse(string urlService, OperationContext opContext, LabourEmployee labourEmployee)
        {
            var proxyLt = new LabourCostingTransService.LabourCostingTransService { Url = urlService + "/LabourCostingTrans" };

            var requestLt = new LabourCostingTransDTO
            {
                transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)), int.Parse(labourEmployee.TransactionDate.Substring(4, 2)), int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                transactionDateSpecified = true,
                labourCostingHours = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingHours) ? Convert.ToDecimal(labourEmployee.LabourCostingHours) : default(decimal),
                labourCostingHoursSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingHours),
                labourCostingValue = !string.IsNullOrWhiteSpace(labourEmployee.LabourCostingValue) ? Convert.ToDecimal(labourEmployee.LabourCostingValue) : default(decimal),
                labourCostingValueSpecified = !string.IsNullOrEmpty(labourEmployee.LabourCostingValue),
                interDistrictCode = labourEmployee.InterDistrictCode,
                accountCode = labourEmployee.AccountCode,
                postingStatus = labourEmployee.PostingStatus,
                project = labourEmployee.Project,
                workOrder = labourEmployee.WorkOrder,
                workOrderTask = labourEmployee.WorkOrderTask,
                employee = labourEmployee.Employee,
                equipmentNo = labourEmployee.EquipmentNo,
                equipmentReference = labourEmployee.EquipmentRef,
                percentComplete = !string.IsNullOrWhiteSpace(labourEmployee.PercentComplete) ? Convert.ToDecimal(labourEmployee.PercentComplete) : default(decimal),
                percentCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.PercentComplete),
                earnCode = labourEmployee.EarnCode,
                labourClass = labourEmployee.LabourClass,
                overtimeInd = labourEmployee.OvertimeInd,
                overtimeIndSpecified = true,
                unitsComplete = !string.IsNullOrWhiteSpace(labourEmployee.UnitsComplete) ? Convert.ToDecimal(labourEmployee.UnitsComplete) : default(decimal),
                unitsCompleteSpecified = !string.IsNullOrEmpty(labourEmployee.UnitsComplete),
                completedCode = labourEmployee.CompletedCode
            };

            //Search Existing
            var requestSearch = new LabourCostingTransSearchParam
            {
                transactionDate = new DateTime(int.Parse(labourEmployee.TransactionDate.Substring(0, 4)),
                    int.Parse(labourEmployee.TransactionDate.Substring(4, 2)),
                    int.Parse(labourEmployee.TransactionDate.Substring(6, 2))),
                transactionDateSpecified = true,
                employee = labourEmployee.Employee,
                project = labourEmployee.Project,
                workOrder = labourEmployee.WorkOrder,
                workOrderTask = labourEmployee.WorkOrderTask
            };
            var searchRestartDto = new LabourCostingTransDTO();
            //La búsqueda solo toma en cuenta la fecha de transacción y el employee id
            var replySearch = proxyLt.search(opContext, requestSearch, searchRestartDto);
            //Existe un elemento
            if (replySearch == null || replySearch.Length < 1) return proxyLt.create(opContext, requestLt);
            foreach (var replyItem in replySearch)
            {
                //Las comparaciones deben hacerse con LPAD para poder establecer bien las comparaciones númericas que trae ellipse en su información
                var equalTranDate =
                    replyItem.labourCostingTransDTO.transactionDate.Equals(requestSearch.transactionDate);
                var equalEmployee = replyItem.labourCostingTransDTO.employee.PadLeft(20, '0')
                    .Equals(requestSearch.employee.PadLeft(20, '0'));
                //posibles nulos de WorkOrders y/o Projects
                var equalWo = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrder) && replyItem
                                    .labourCostingTransDTO.workOrder.PadLeft(20, '0')
                                    .Equals(requestSearch.workOrder.PadLeft(20, '0'));

                string itemTaskNo = string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.workOrderTask) ? "000" : replyItem.labourCostingTransDTO.workOrderTask;
                var equalTask = itemTaskNo.Equals(!string.IsNullOrWhiteSpace(requestSearch.workOrderTask) ? requestSearch.workOrderTask.PadLeft(3, '0') : "000");
                var equalProject = !string.IsNullOrWhiteSpace(replyItem.labourCostingTransDTO.project) &&
                                    replyItem.labourCostingTransDTO.project.PadLeft(20, '0')
                                        .Equals(requestSearch.project.PadLeft(20, '0'));

                if (!equalTranDate || !equalEmployee || ((!equalWo || !equalTask) && !equalProject)) continue;
                var result = proxyLt.delete(opContext, replySearch[0].labourCostingTransDTO);
                return result;
            }
            return null;
        }

        /// <summary>
        /// Carga un registro de labor para el empleado asignado
        /// </summary>
        /// <param name="opSheet"></param>
        /// <param name="labourEmployee"></param>
        /// <param name="replaceExisting">bool: Si es true y ya existe un registro con la misma ot-tarea se modificará por el nuevo registro. Si es false, se ignorará el registro y siempre se adicionará uno nuevo</param>
        [Obsolete("Not used anymore", true)]//Deprecated
        public void LoadEmployeeMso(Screen.OperationContext opSheet, LabourEmployee labourEmployee, bool replaceExisting = true)
        {
            //Proceso del screen
            var proxySheet = new Screen.ScreenService();

            var requestSheet = new Screen.ScreenSubmitRequestDTO();
            var arrayFields = new ArrayScreenNameValue();

            //Selección de ambiente
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            //Aseguro que no esté en alguna pantalla antigua
            _eFunctions.RevertOperation(opSheet, proxySheet);
            //ejecutamos el programa
            var replySheet = proxySheet.executeScreen(opSheet, "MSO850");

            //validamos el ingreso al programa
            if (replySheet == null || replySheet.mapName != "MSM850A" || ValidateError(replySheet) || _eFunctions.CheckReplyWarning(replySheet))
                throw new Exception("No se pudo establecer comunicación con el servicio");
            //Enviamos datos principales para activar los campos de labor
            arrayFields.Add("EMP_ID1I", labourEmployee.Employee);
            arrayFields.Add("TRAN_DATE1I", labourEmployee.TransactionDate);

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";

            replySheet = proxySheet.submit(opSheet, requestSheet);

            //Continuamos en la pantalla pero con los campos de labor activos
            if (!ValidateError(replySheet) && replySheet.mapName == "MSM850A")
            {
                var labourFoundFlag = false;
                var rowMso = 1;
                //obtenemos los campos del reply
                var replyFields = new ArrayScreenNameValue(replySheet.screenFields);
                //son variables para determinar el cambio de screen real
                //reajustamos el valor de la tarea a un numérico ###
                if (string.IsNullOrWhiteSpace(labourEmployee.WorkOrderTask) && cbAutoTaskAssigment.Checked)
                    if (!string.IsNullOrWhiteSpace(labourEmployee.WorkOrder))
                        labourEmployee.WorkOrderTask = labourEmployee.WorkOrderTask == "" ? "" : "001";

                //iniciamos el recorrido
                while (!labourFoundFlag)
                {
                    //comprobamos que: 1. Que no exista OT-Tarea cargada para que no se duplique, y 2. nos ubicamos en el último campo disponible
                    //si existe 1, entonces actualizamos la información. Si no, continuamos a ubicarnos en 2.
                    var isEmpty = replyFields.GetField("WO_PROJ1I" + rowMso).value.Equals("");
                    var sameWo = labourEmployee.WorkOrder.Equals(replyFields.GetField("WO_PROJ1I" + rowMso).value);
                    var screenTaskValue = replyFields.GetField("TASK1I" + rowMso).value;
                    var isWopInd = replyFields.GetField("WOP_IND1I" + rowMso).value.Equals("W");

                    if (string.IsNullOrWhiteSpace(screenTaskValue) && isWopInd && cbAutoTaskAssigment.Checked)
                    {
                        screenTaskValue = "001";
                    }

                    var sameTask = labourEmployee.WorkOrderTask == screenTaskValue;
                    var sameEarnClass = string.IsNullOrWhiteSpace(labourEmployee.EarnCode) || labourEmployee.EarnCode.Equals(replyFields.GetField("EARN_CLASS1I" + rowMso).value);
                    //si se encuentra una posición para escribir la labor se activa el flag y se sale del while
                    if (isEmpty || (replaceExisting && sameWo && sameTask && sameEarnClass))
                    {
                        labourFoundFlag = true;
                        continue;
                    }
                    rowMso++;
                    //si es mayor que los registros del screen (5) entonces envíe el screen y prosiga con los del siguiente screen
                    if (rowMso <= 5) continue;

                    var previousReply = replyFields;

                    var errorLoopPos = 0;//para controlar que no se quede atrapado en loop infinito
                    //verifico la respuesta de este envío por errores y advertencias
                    while (replySheet.mapName == "MSM850A" && replyFields.GetField("TRAN_DATE1I").value == labourEmployee.TransactionDate && !ValidateError(replySheet))
                    {
                        //Creamos la nueva acción de envío reutilizando los elementos anteriores
                        requestSheet = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                        replySheet = proxySheet.submit(opSheet, requestSheet);
                        replyFields = new ArrayScreenNameValue(replySheet.screenFields);

                        //para asegurar que hubo un cambio real de pantalla. De lo contrario pasará hasta el final de los envíos
                        var currentReply = replyFields;
                        if (previousReply == currentReply)
                        {
                            errorLoopPos++;
                            if (errorLoopPos > 50)
                                throw new Exception("Error de ejecución. El proceso ha alcanzo el límite de intentos en posicionamiento");
                            continue;
                        }
                        break;
                    }
                    if (ValidateError(replySheet) || _eFunctions.CheckReplyWarning(replySheet))
                        throw new Exception(replySheet.message);

                    rowMso = 1;
                }
                //ingresamos los elementos para los campos a enviar   
                arrayFields.Add("ORD_HRS1I" + rowMso, labourEmployee.LabourCostingHours);
                arrayFields.Add("OT_HRS1I" + rowMso, labourEmployee.OvertimeInd ? labourEmployee.LabourCostingHours : null);
                arrayFields.Add("ACCOUNT1I" + rowMso, labourEmployee.AccountCode);
                arrayFields.Add("WOP_IND1I" + rowMso, string.IsNullOrWhiteSpace(labourEmployee.WorkOrder) ? "P" : "W");
                arrayFields.Add("WO_PROJ1I" + rowMso, string.IsNullOrWhiteSpace(labourEmployee.WorkOrder) ? labourEmployee.Project : labourEmployee.WorkOrder);
                arrayFields.Add("TASK1I" + rowMso, labourEmployee.WorkOrderTask);
                arrayFields.Add("EQUIPMENT1I" + rowMso, labourEmployee.EquipmentNo);
                arrayFields.Add("UNITS_COMP1I" + rowMso, labourEmployee.UnitsComplete);
                arrayFields.Add("PC_COMP1I" + rowMso, labourEmployee.PercentComplete);
                arrayFields.Add("CODE_COMP1I" + rowMso, labourEmployee.CompletedCode);
                arrayFields.Add("EARN_CLASS1I" + rowMso, labourEmployee.EarnCode);
                arrayFields.Add("LABOUR_CLASS1I" + rowMso, labourEmployee.LabourClass);
                //enviamos la información
                requestSheet = new Screen.ScreenSubmitRequestDTO
                {
                    screenFields = arrayFields.ToArray(),
                    screenKey = "1"
                };
                replySheet = proxySheet.submit(opSheet, requestSheet);

                var errorLoopVal = 0;

                while (replySheet != null && !ValidateError(replySheet) &&//no existen errores
                       (_eFunctions.CheckReplyWarning(replySheet) || replySheet.functionKeys == "XMIT-Confirm" || replySheet.functionKeys == "XMIT-Validate"))//requiere confirmación
                {
                    errorLoopVal++;
                    if (errorLoopVal > 50)
                        throw new Exception("Error de ejecución. El proceso ha alcanzo el límite de intentos en confirmación");
                    requestSheet = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                    replySheet = proxySheet.submit(opSheet, requestSheet);
                }

                if (replySheet != null && replySheet.message.Length > 2 && replySheet.message.Substring(0, 2) == "X2")
                    throw new Exception(replySheet.message);

                return;
            }

            if (replySheet == null)
                throw new Exception("No se puede establecer conexión con el programa Mso850");

            throw new Exception(replySheet.message);
        }

        /// <summary>
        /// Verificar Error en Reply de Screen. Solo aplica para Mso850 por el comportamiento particular para algunos mensajes específicos
        /// </summary>
        /// <param name="reply"></param>
        /// <returns></returns>
        public bool ValidateError(Screen.ScreenDTO reply)
        {
            //Si no existe un reply es error de ejecución. O si el reply tiene un error de datos
            if (reply == null)
            {
                Debugger.LogError("RibbonEllipse:ValidateError(Screen.ScreenDTO)", "null reply error");
                return true;
            }
            if (reply.message.Length < 2 || reply.message.Substring(0, 2) != "X2") return false;
            //0039:WO NOT ON FILE
            //7630 WO TASK NOT ON FILE
            //8438 HRS NORMALES INTRODUCIR PARA COD GANANCIAS TIPO OT
            //5331 WORK ORDER IS CLOSED TO COMMITMENT
            if (reply.message.Substring(3, 4).Equals("6008")) //6008:WARNING - VALUE IS ZERO (ORD_HRS); 
                return false;
            if (reply.message.Substring(3, 4).Equals("8852")) //8852 UNPROCESSED COSTING DETAILS DISPLAYED
                return false;
            if (reply.message.Substring(3, 4).Equals("4744")) //4744
                return false;
            if (reply.message.Substring(3, 4).Equals("8839")) //8839
                return false;

            Debugger.LogError("LabourCosting.ValidateError(ScreenDTO reply)", "Error: " + reply.message);
            return true;
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

        public static class LabourSheetTypeConstants
        {
            public static string Default = "Default";
            public static string Mso850 = "MSO850";//Deprecated
            public static string Mse850 = "MSE850";
            public static string Elecsa = "ELECSA";
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
                    _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                    var district = "" + _cells.GetCell("B3").Value;
                    for (var i = 0; i < OtFields; i++)
                    {
                        var target = _cells.GetCell(DefaultStartCol + i * NroEle, DefaultTitleRow - 2);
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
                    _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                    var district = "" + _cells.GetCell("B3").Value;
                    var i = ElecsaTitleRow + 1;
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
    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class LabourEmployee
    {
        public string TransactionDate;
        public string LabourCostingHours;
        public bool OvertimeInd;
        public string LabourCostingValue;
        public string InterDistrictCode;
        public string AccountCode;
        public string PostingStatus;
        public string Project;
        public string WorkOrder;
        public string WorkOrderTask;
        public string Employee;
        public string EquipmentNo;
        public string EquipmentRef;
        public string UnitsComplete;
        public string PercentComplete;
        public string CompletedCode;
        public string EarnCode;
        public string LabourClass;


    }
    public static class Queries
    {
        public static string GetGroupEmployeesQuery(string workGroup, string dbReference, string dbLink)
        {
            var query = "SELECT "+
                        "     EMP.EMPLOYEE_ID   CEDULA, "+
                        "     EMP.FIRST_NAME || ' ' || EMP.SURNAME NOMBRE "+
                        " FROM "+
                        "     " + dbReference + ".MSF723" + dbLink + " WE "+
                        "     INNER JOIN " + dbReference + ".MSF810" + dbLink + " EMP "+
                        "     ON EMP.EMPLOYEE_ID   = WE.EMPLOYEE_ID "+
                        "     OR TRIM(EMP.PREF_NAME)   = TRIM(WE.EMPLOYEE_ID) " +
                        " WHERE " +
                        "     WE.WORK_GROUP = '" + workGroup + "' "+
                        "     AND ( WE.STOP_DT_REVSD   = '00000000' "+
                        "           OR ( 99999999 - WE.STOP_DT_REVSD ) >= TO_CHAR( SYSDATE, 'YYYYMMDD' ) ) "+
                        "     AND WE.REC_723_TYPE    = 'W'";

            query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            return query;
        }
    }
}
