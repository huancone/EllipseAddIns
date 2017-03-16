using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using System.Web.Services.Ellipse;
using EllipseWorkOrdersClassLibrary;
using EllipseStdTextClassLibrary;
using System.Threading;

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
        const int Mso850TitleRow = 5;
        const int ElecsaTitleRow = 6;
        const int ElecsaTitleColumn = 6;
        private const int Mso850ResultColumn = 18;
        private const int ElecsaResultColumn = 30;
        private const string TableNameMso850 = "Mso850Table";
        private const string TableNameDefault = "LabourDefaultTable";
        private const string TableNameElecsa = "ElecsaTable";

        private const int OtFields = 20;

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

        private void btnFormatHeader_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetHeaderData();
        }

        private void btnFormatDefault_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
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
            FormatMso850Sheet();
        }

        private void btnFormatElecsa_Click(object sender, RibbonControlEventArgs e)
        {
            FormatElecsaSheet();
        }

        private void btnLoadLaborSheet_Click(object sender, RibbonControlEventArgs e)
        {
			//si ya hay un thread corriendo que no se ha detenido
			if (_thread != null && _thread.IsAlive) return;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            if (!_cells.IsDecimalDotSeparator())
                if (MessageBox.Show(@"El separador de decimales configurado actualmente no es el punto. Usar un separador de decimales diferente puede generar errores al momento de cargar valores numéricos. ¿Está seguro que desea continuar?",  @"ALERTA DE SEPARADOR DE DECIMALES", MessageBoxButtons.OKCancel) != DialogResult.OK) return;

            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Default))
                _thread = new Thread(LoadDefaultLabourCost);
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mso850))
                _thread = new Thread(LoadMso850LabourCost);
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
                Debugger.LogError("RibbonEllipse:FormatSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
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

                var sqlQuery = Queries.GetGroupEmployeesQuery(groupName, _eFunctions.dbReference, _eFunctions.dbLink);

                if (_eFunctions.DebugQueries)
                    _cells.GetCell("L1").Value = sqlQuery;

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
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
                    _cells.GetCell(DefaultStartCol + i*NroEle, DefaultTitleRow - 3).Value = "OT " + (i + 1);
                    _cells.GetCell(DefaultStartCol + i*NroEle, DefaultTitleRow - 3).Style =
                        _cells.GetStyle(StyleConstants.TitleOptional);
                    _cells.MergeCells(DefaultStartCol + i*NroEle, DefaultTitleRow - 3,
                        DefaultStartCol + i*NroEle + (NroEle - 1), DefaultTitleRow - 3);
                    //Número OT
                    _cells.GetCell(DefaultStartCol + i*NroEle, DefaultTitleRow - 2).Style =
                        _cells.GetStyle(StyleConstants.Select);
                    vstoSheet.Controls.Remove("SeekOrder" + i);
                    var orderNameRange =
                        vstoSheet.Controls.AddNamedRange(
                            _cells.GetCell(DefaultStartCol + i*NroEle, DefaultTitleRow - 2),
                            "SeekOrder" + i);
                    orderNameRange.Change += GetDefaultWorkOrderDescriptionChangedValue;
                    _cells.GetRange(DefaultStartCol + i*NroEle, DefaultTitleRow - 2,
                        DefaultStartCol + i*NroEle + (NroEle - 1), DefaultTitleRow - 2)
                        .NumberFormat = NumberFormatConstants.Text;
                    _cells.MergeCells(DefaultStartCol + i*NroEle, DefaultTitleRow - 2,
                        DefaultStartCol + i*NroEle + (NroEle - 1), DefaultTitleRow - 2);
                    //Descripción OT
                    _cells.GetCell(DefaultStartCol + i*NroEle, DefaultTitleRow - 1).Style =
                        _cells.GetStyle(StyleConstants.Option);
                    _cells.MergeCells(DefaultStartCol + i*NroEle, DefaultTitleRow - 1,
                        DefaultStartCol + i*NroEle + (NroEle - 1), DefaultTitleRow - 1);
                    //Cada componente
                    _cells.GetCell(DefaultStartCol + i*NroEle + 0, DefaultTitleRow).Value = "HRS_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i*NroEle + 0, DefaultTitleRow).AddComment("hh.mm");
                    _cells.GetCell(DefaultStartCol + i*NroEle + 0, DefaultTitleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleRequired);
                    _cells.GetCell(DefaultStartCol + i*NroEle + 1, DefaultTitleRow).Value = "TASK_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i*NroEle + 1, DefaultTitleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleRequired);
                    _cells.GetCell(DefaultStartCol + i*NroEle + 2, DefaultTitleRow).Value = "ECODE_" + (i + 1);
                    _cells.GetCell(DefaultStartCol + i*NroEle + 2, DefaultTitleRow).Style =
                        _cells.GetStyle(StyleConstants.TitleOptional);

                }
                _cells.GetRange(DefaultStartCol, DefaultTitleRow + 1, DefaultStartCol + OtFields*NroEle,
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
                    _cells.GetCell(DefaultStartCol + j*NroEle + 2, DefaultTitleRow + 1).Validation.Delete();
                    _cells.GetCell(DefaultStartCol + j*NroEle + 2, DefaultTitleRow + 1)
                        .Validation.Add((Excel.XlDVType) validation.Type, validation.AlertStyle, validation.Operator,
                            validation.Formula1, validation.Formula2);
                }
                _cells.FormatAsTable(
                    _cells.GetRange(1, DefaultTitleRow, DefaultStartCol + OtFields*NroEle - 1, DefaultTitleRow + 1),
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
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace,
                    _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar obtener el formato del grupo seleccionado");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
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

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

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
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
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
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
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
                _cells.SetValidationList(_cells.GetCell(2, ElecsaTitleRow - 3), DistrictConstants.GetDistrictList(), ValidationSheetName, 1);
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
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        public void LoadDefaultLabourCost()
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01 + LabourSheetTypeConstants.Default)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnviroment.Label == null || drpEnviroment.Label.Equals(""))
                    throw new ArgumentException("Seleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return; //no se autentica, se cancela el proceso



                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = _eFunctions.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                var i = 10;
                //recorro la lista de empleados de forma vertical
                while ("" + _cells.GetCell(2, i).Value != "")
                {
                    //proceso por empleado
                    var employee = new LabourEmployee
                    {
                        WOP_IND = "W",
                        EMP_ID = ("" + _cells.GetCell(2, i).Value).Trim(),
                        TRAN_DATE = ("" + _cells.GetCell(2, 5).Value).Trim()
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


                                employee.ORD_HRS = _cells.GetEmptyIfNull(_cells.GetCell(j, i).Value2);
                                employee.TASK = _cells.GetEmptyIfNull(_cells.GetCell(j + 1, i).Value2);
                                if (string.IsNullOrWhiteSpace(employee.TASK))
                                    employee.TASK = "001";
                                //el número de tarea debe tener tres dígitos 001
                                if (employee.TASK.Length > 1 && employee.TASK.Length < 3)
                                    employee.TASK = employee.TASK.PadLeft(3, '0');
                                employee.EARN_CLASS = earningClass;

                                employee.WO_PROJ = _cells.GetEmptyIfNull(_cells.GetCell(j, DefaultTitleRow - 2).Value2);

                                LoadEmployee(opSheet, employee);

                                _cells.GetRange(j, i, j, i).ClearComments();
                                _cells.GetRange(j, i, j + NroEle - 1, i).Style = StyleConstants.Success;
                                _cells.GetRange(j, i, j, i).Select();
                            }
                        }
                        catch (Exception ex)
                        {
                            Debugger.LogError("RibbonEllipse:LoadDefaultLabourCost()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace, _eFunctions.DebugErrors);
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
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace,
                    _eFunctions.DebugErrors);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }

        }

        public void LoadMso850LabourCost()
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01 + LabourSheetTypeConstants.Mso850)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnviroment.Label == null || drpEnviroment.Label.Equals(""))
                    throw new ArgumentException("Seleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return; //no se autentica, se cancela el proceso



                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = _eFunctions.DebugWarnings
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
                            TRAN_DATE = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                            ORD_HRS = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value),
                            OT_HRS = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value),
                            VALUE = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value),
                            INT_DSTRCT_CDE = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value),
                            ACCOUNT = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value),
                            STATUS = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value),
                            WOP_IND = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value),
                            WO_PROJ = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value),
                            TASK = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                            EMP_ID = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value),
                            EQUIPMENT = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value),
                            UNITS_COMP = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value),
                            PC_COMP = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value),
                            CODE_COMP = codeComp,
                            EARN_CLASS = earningCode,
                            LABOUR_CLASS = laborClass
                        };

                        if (string.IsNullOrWhiteSpace(employee.TASK))
                            employee.TASK = "001";
                        if (employee.TASK.Length > 1 && employee.TASK.Length < 3)
                            employee.TASK = employee.TASK.PadLeft(3, '0');

                        LoadEmployee(opSheet, employee);
                        _cells.GetCell(Mso850ResultColumn, i).Value = "SUCCESS";
                        _cells.GetCell(Mso850ResultColumn, i).Style = StyleConstants.Success;
                        _cells.GetCell(Mso850ResultColumn, i).Select();
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:LoadMso850LabourCost()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace,
                            _eFunctions.DebugErrors);
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
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace,
                    _eFunctions.DebugErrors);
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
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01 + LabourSheetTypeConstants.Elecsa)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");

                if (drpEnviroment.Label == null || drpEnviroment.Label.Equals(""))
                    throw new ArgumentException("Seleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return; //no se autentica, se cancela el proceso



                //Se usa un solo OperationContext para ahorrar en recursos y solicitudes
                var opSheet = new Screen.OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    returnWarnings = _eFunctions.DebugWarnings
                };

                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                var districtCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, ElecsaTitleRow - 3).Value2);
                var loadType = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, ElecsaTitleRow - 3).Value2);

                var i = ElecsaTitleRow + 1;
                const string employeeId = "CONPBVELE";
                const string wpFlag = "W";
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
                        task = task ?? "001";
                        //el número de tarea debe tener tres dígitos 001
                        if (task.Length > 1 && task.Length < 3)
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
                                    TRAN_DATE = transDate,
                                    ORD_HRS = hours,
                                    ACCOUNT = costCenter,
                                    WOP_IND = wpFlag,
                                    WO_PROJ = workOrder,
                                    TASK = task,
                                    EMP_ID = employeeId,
                                    EARN_CLASS = earningCode,
                                    LABOUR_CLASS = laborClass
                                };

                                LoadEmployee(opSheet, employee, false);
                                _cells.GetCell(j, i).ClearComments();
                                _cells.GetCell(j, i).Style = StyleConstants.Success;
                                _cells.GetCell(j, i).Select();
                            }
                            catch (Exception ex)
                            {
                                errorFlag = true;
                                Debugger.LogError("RibbonEllipse:LoadElecsaLabourCost()",
                                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                    ex.StackTrace,
                                    _eFunctions.DebugErrors);
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
                    returnWarnings = _eFunctions.DebugWarnings
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                //recorro la lista de tareas de forma vertical para duración
                while (
                    !(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2).Equals("") &&
                      _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals("")))
                {
                    try
                    {
                        var workOrder = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value2);
                        var durationCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 1, i).Value2);
                        var startHour = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 2, i).Value2);
                        var endHour = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 3, i).Value2);
                        var wo = WorkOrderActions.GetNewWorkOrderDto(workOrder);
                        var completeCommentToAppend =
                            _cells.GetNullIfTrimmedEmpty(_cells.GetCell(ElecsaResultColumn + 4, i).Value2);

                        var duration = new WorkOrderDuration
                        {
                            jobDurationsDate = transDate,
                            jobDurationsCode = durationCode,
                            jobDurationsStart = startHour,
                            jobDurationsFinish = endHour
                        };


                        var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                        WorkOrderActions.CreateWorkOrderDuration(urlService, opWo, districtCode, wo, duration);

                        var stdTextId = "CW" + districtCode + workOrder;

                        var stdTextCopc = StdText.GetCustomOpContext(opWo.district, opWo.position, opWo.maxInstances,
                            opWo.returnWarnings);
                        var woCompleteComment = StdText.GetCustomText(urlService, stdTextCopc, stdTextId);

                        StdText.SetCustomText(urlService, stdTextCopc, stdTextId,
                            woCompleteComment + "\n" + completeCommentToAppend);

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
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace,
                    _eFunctions.DebugErrors);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void CleanLabourTable()
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Mso850))
                _cells.ClearTableRange(TableNameMso850);
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Elecsa))
                _cells.ClearTableRange(TableNameElecsa);
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01 + LabourSheetTypeConstants.Default))
                _cells.ClearTableRange(TableNameDefault);

        }

        /// <summary>
        /// Carga un registro de labor para el empleado asignado
        /// </summary>
        /// <param name="opSheet"></param>
        /// <param name="labourEmployee"></param>
        /// <param name="replaceExisting">bool: Si es true y ya existe un registro con la misma ot-tarea se modificará por el nuevo registro. Si es false, se ignorará el registro y siempre se adicionará uno nuevo</param>
        public void LoadEmployee(Screen.OperationContext opSheet, LabourEmployee labourEmployee, bool replaceExisting = true)
        {
            //Proceso del screen
            var proxySheet = new Screen.ScreenService();

            var requestSheet = new Screen.ScreenSubmitRequestDTO();
            var arrayFields = new ArrayScreenNameValue();

            //Selección de ambiente
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
            //Aseguro que no esté en alguna pantalla antigua
            _eFunctions.RevertOperation(opSheet, proxySheet);
            //ejecutamos el programa
            var replySheet = proxySheet.executeScreen(opSheet, "MSO850");

            //validamos el ingreso al programa
            if (replySheet == null || replySheet.mapName != "MSM850A" || ValidateError(replySheet) || _eFunctions.CheckReplyWarning(replySheet))
                throw new Exception("No se pudo establecer comunicación con el servicio");
            //Enviamos datos principales para activar los campos de labor
            arrayFields.Add("EMP_ID1I", labourEmployee.EMP_ID);
            arrayFields.Add("TRAN_DATE1I", labourEmployee.TRAN_DATE);

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
                if (string.IsNullOrWhiteSpace(labourEmployee.TASK))
                    labourEmployee.TASK = "001";

                //iniciamos el recorrido
                while (!labourFoundFlag)
                {
                    //comprobamos que: 1. Que no exista OT-Tarea cargada para que no se duplique, y 2. nos ubicamos en el último campo disponible
                    //si existe 1, entonces actualizamos la información. Si no, continuamos a ubicarnos en 2.
                    var isEmpty = replyFields.GetField("WO_PROJ1I" + rowMso).value.Equals("");
                    var sameWo = labourEmployee.WO_PROJ.Equals(replyFields.GetField("WO_PROJ1I" + rowMso).value);
                    var screenTaskValue = replyFields.GetField("TASK1I" + rowMso).value;
                    if (string.IsNullOrWhiteSpace(screenTaskValue))
                        screenTaskValue = "001";
                    var sameTask = int.Parse(labourEmployee.TASK) == int.Parse(screenTaskValue);
                    var sameEarnClass = string.IsNullOrWhiteSpace(labourEmployee.EARN_CLASS) || labourEmployee.EARN_CLASS.Equals(replyFields.GetField("EARN_CLASS1I" + rowMso).value);
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
                    while (replySheet.mapName == "MSM850A" && replyFields.GetField("TRAN_DATE1I").value == labourEmployee.TRAN_DATE && !ValidateError(replySheet))
                    {
                        //Creamos la nueva acción de envío reutilizando los elementos anteriores
                        requestSheet = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                        replySheet = proxySheet.submit(opSheet, requestSheet);
                        replyFields = new ArrayScreenNameValue(replySheet.screenFields);

                        //para asegurar que hubo un cambio real de pantalla. De lo contrario pasará hasta el final de los envíos
                        var currentReply = replyFields;
                        if(previousReply == currentReply)
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
                arrayFields.Add("ORD_HRS1I" + rowMso, labourEmployee.ORD_HRS);
                arrayFields.Add("OT_HRS1I" + rowMso, labourEmployee.OT_HRS);
                arrayFields.Add("ACCOUNT1I" + rowMso, labourEmployee.ACCOUNT);
                arrayFields.Add("WOP_IND1I" + rowMso, labourEmployee.WOP_IND);
                arrayFields.Add("WO_PROJ1I" + rowMso, labourEmployee.WO_PROJ);
                arrayFields.Add("TASK1I" + rowMso, labourEmployee.TASK);
                arrayFields.Add("EQUIPMENT1I" + rowMso, labourEmployee.EQUIPMENT);
                arrayFields.Add("UNITS_COMP1I" + rowMso, labourEmployee.UNITS_COMP);
                arrayFields.Add("PC_COMP1I" + rowMso, labourEmployee.PC_COMP);
                arrayFields.Add("CODE_COMP1I" + rowMso, labourEmployee.CODE_COMP);
                arrayFields.Add("EARN_CLASS1I" + rowMso, labourEmployee.EARN_CLASS);
                arrayFields.Add("LABOUR_CLASS1I" + rowMso, labourEmployee.LABOUR_CLASS);
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
                    if(errorLoopVal>50)
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
                Debugger.LogError("RibbonEllipse:ValidateError(Screen.ScreenDTO)", "null reply error",
                    _eFunctions.DebugErrors);
                if (_eFunctions.DebugErrors)
                    MessageBox.Show(@"Se ha producido un error en tiempo de ejecución");
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

            if (_eFunctions.DebugErrors)
                MessageBox.Show(@"Error: " + reply.message);
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
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

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
                        _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

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
            public static string Mso850 = "MSO850";
            public static string Elecsa = "ELECSA";
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null &&_thread.IsAlive)
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
                    _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                    var district = "" + _cells.GetCell("B3").Value;
                    for (var i = 0; i < OtFields; i++)
                    {
                        var target = _cells.GetCell(DefaultStartCol + i*NroEle, DefaultTitleRow - 2);
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
                    _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
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
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
    }

    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class LabourEmployee
    {
        public string TRAN_DATE;
        public string ORD_HRS;
        public string OT_HRS;
        public string VALUE;
        public string INT_DSTRCT_CDE;
        public string ACCOUNT;
        public string STATUS;
        public string WOP_IND;
        public string WO_PROJ;
        public string TASK;
        public string EMP_ID;
        public string EQUIPMENT;
        public string UNITS_COMP;
        public string PC_COMP;
        public string CODE_COMP;
        public string EARN_CLASS;
        public string LABOUR_CLASS;
    }
    public static class Queries
    {
        public static string GetGroupEmployeesQuery(string workGroup, string dbReference, string dbLink)
        {
            var sqlQuery = " SELECT DISTINCT (TRIM(EMP.SURNAME) || ' ' || TRIM(EMP.FIRST_NAME)) NOMBRE ,GEMP.EMPLOYEE_ID CEDULA" +
            " FROM " + dbReference + ".MSF723" + dbLink + " GEMP" +
            "     JOIN " + dbReference + ".MSF810" + dbLink + " EMP ON GEMP.EMPLOYEE_ID = EMP.EMPLOYEE_ID" +
            "     JOIN " + dbReference + ".MSF826" + dbLink + " LEMP ON GEMP.EMPLOYEE_ID = LEMP.EMPLOYEE_ID" +
            "     JOIN " + dbReference + ".MSF855" + dbLink + " LCOST ON LEMP.LAB_COST_CLASS = LCOST.LAB_COST_CLASS" +
            " WHERE GEMP.WORK_GROUP = '" + workGroup + "'" +
            "     AND (GEMP.STOP_DT_REVSD  = '00000000' OR (99999999 - GEMP.STOP_DT_REVSD) >= TO_CHAR(SYSDATE, 'YYYYMMDD'))" +
            "     AND GEMP.REC_723_TYPE = 'W'" +
            " ORDER BY CEDULA";

            return sqlQuery;
        }
    }
}
