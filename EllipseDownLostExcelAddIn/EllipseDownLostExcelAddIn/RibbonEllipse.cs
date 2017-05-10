using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseEquipmentClassLibrary;
using EllipseStdTextClassLibrary;
using EllipseWorkOrdersClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Screen = EllipseCommonsClassLibrary.ScreenService;

namespace EllipseDownLostExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;
        private const string SheetName01 = "DownLostSheet";
        private const string SheetName02 = "DownCodeList";
        private const string SheetName03 = "LostCodeList";
        private const string SheetName04 = "GeneratedCollection";
        private const string SheetNameP01 = "DownLostSheetPBV";
        private const string ValidationSheetName = "ValidationSheet";

        private const int TitleRow01 = 7;
        private const int ResultColumn01 = 15;
        private const int ResultColumnP01 = 19;

        private const string TableName01 = "DownLostTable";
        private const string TableName02 = "DownCodeTable";
        private const string TableName03 = "LostCodeTable";
        private const string TableName04 = "GeneratedCollectionTable";

        private bool _ignoreDuplicate;
        private Thread _thread;
        private string _woDownOriginator;

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

        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLost()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }
        private void btnReviewDLPbv_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewDownLostPbv);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLostPbv()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnCreateDL_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _ignoreDuplicate = false;
                    _thread = new Thread(CreateDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnCreatIgnoreDuplicate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _ignoreDuplicate = true;
                    _thread = new Thread(CreateDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnDeleteDL_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    var dr = MessageBox.Show(@"Esta acción eliminará los registros Down/Lost existentes. ¿Está seguro que desea continuar?", @"ELIMINAR DOWN Y LOST", MessageBoxButtons.YesNo);
                    if (dr != DialogResult.Yes)
                        return;
                    _thread = new Thread(DeleteDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DeleteDownLost()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnClearTable_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
        }

        private void btnFormatDownPbv_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetPbv();
        }

        private void btnGenerateCollection_Click(object sender, RibbonControlEventArgs e)
        {
            GenerateCollectionList();
        }


        /// <summary>
        /// Formatea la hoja con la estructura y estilo que se requiere
        /// </summary>
        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(SheetName04);
                _cells.CreateNewWorksheet(ValidationSheetName);

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "DOWN TIME & LOST PRODUCTION - ELLIPSE 8";
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
                //Cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                //Cells.GetCell("K5").Style = Cells.GetStyle(StyleConstants.TitleAdditional);


                _cells.GetCell("A3").Value = "DISTRICT";
                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName, 1);
                _cells.GetCell("B3").Value = "ICOR";

                var equipTypeList = new List<string> {"EQUIPMENT", "EGI", "LIST TYPE", "PROD.UNIT"};

                _cells.SetValidationList(_cells.GetCell("A4"), equipTypeList, ValidationSheetName, 2);
                _cells.GetCell("A4").Value = "EQUIPMENT";

                _cells.GetCell("A5").Value = "LIST ID";
                _cells.GetCell("A5").AddComment("Solo para List Type");

                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("D4").AddComment("yyyyMMdd");

                _cells.GetCell("C5").Value = "TYPE";
                var dataTypeList = new List<string> {"DOWN", "LOST", "DOWN & LOST"};

                _cells.SetValidationList(_cells.GetCell("D5"), dataTypeList, ValidationSheetName, 3);
                _cells.GetCell("D5").Value = "DOWN & LOST";
                //Valores predeterminados
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);


                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetRange("A3", "D5").NumberFormat = NumberFormatConstants.Text;

                //GENERAL

                _cells.GetCell(01, TitleRow01).Value = "EQUIP_NO";
                _cells.GetCell(02, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(02, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(03, TitleRow01).Value = "COMP_MOD_CODE";
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(04, TitleRow01).Value = "START_DATE";
                _cells.GetCell(04, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(05, TitleRow01).Value = "START_TIME";
                _cells.GetCell(05, TitleRow01).AddComment("hhmm");
                _cells.GetCell(06, TitleRow01).Value = "FINISH_DATE";
                _cells.GetCell(06, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(06, TitleRow01).AddComment("yyyyMMdd - Requerido si se usa COLLECTION");
                _cells.GetCell(07, TitleRow01).Value = "FINISH_TIME";
                _cells.GetCell(07, TitleRow01).AddComment("hhmm");
                _cells.GetCell(08, TitleRow01).Value = "ELAPSED";
                _cells.GetCell(08, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(09, TitleRow01).Value = "COLLECTION";
                _cells.GetCell(09, TitleRow01).AddComment(ShiftConstants.ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftConstants.ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftConstants.ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftConstants.ShiftCodes.DailyMorningCode + ": Dia 06-06");
                _cells.GetCell(09, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow01).Value = "SHIFT";
                _cells.GetCell(10, TitleRow01).AddComment("Este campo será ignorado si usa alguna colección");
                _cells.GetCell(11, TitleRow01).Value = "EVENT_TYPE";
                _cells.GetCell(11, TitleRow01).AddComment("LOST/DOWN");
                _cells.GetCell(12, TitleRow01).Value = "EVENT_CODE";
                _cells.GetCell(13, TitleRow01).Value = "EVENT_DESC";
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(14, TitleRow01).Value = "WORKORDER/COMENTARIO";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(14, TitleRow01).AddComment("WorkOrder para Down, Comentario para Lost");
                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;

                //Adición de validaciones de campo
                var collectionList = new List<string>
                {
                    ShiftConstants.ShiftCodes.HourToHourCode,
                    ShiftConstants.ShiftCodes.DailyZeroCode,
                    ShiftConstants.ShiftCodes.DailyMorningCode,
                    ShiftConstants.ShiftCodes.DayNightCode
                };
                _cells.SetValidationList(_cells.GetRange(09, TitleRow01 + 1, 09, TitleRow01 + 101), collectionList, ValidationSheetName, 4);

                var typeEvent = new List<string> {"DOWN", "LOST"};
                _cells.SetValidationList(_cells.GetRange(11, TitleRow01 + 1, 11, TitleRow01 + 101), typeEvent, ValidationSheetName, 5);


                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 2
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);

                _cells.GetCell("B1").Value = "DOWN TIME CODE LIST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "I2");


                _cells.GetCell("A4").Value = "CÓDIGO";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell("B4").Value = "DESCRIPCIÓN";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetRange("A4", "B4").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.FormatAsTable(_cells.GetRange("A4", "B5"), TableName02);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                var dr = _eFunctions.GetQueryResult(Queries.GetDownTimeCodeListQuery(_eFunctions.dbReference, _eFunctions.dbLink));

                if (dr != null && !dr.IsClosed && dr.HasRows)
                {
                    var i = 5;
                    while (dr.Read())
                    {
                        _cells.GetCell(1, i).Value = "'" + dr["CODE"].ToString().Trim();
                        _cells.GetCell(2, i).Value = dr["DESCRIPTION"].ToString().Trim();
                        i++;
                    }
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();



                //CONSTRUYO LA HOJA 3
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);

                _cells.GetCell("B1").Value = "LOST PRODUCTION CODE LIST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "I2");


                _cells.GetCell("A4").Value = "CÓDIGO";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell("B4").Value = "DESCRIPCIÓN";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetRange("A4", "B4").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.FormatAsTable(_cells.GetRange("A4", "B5"), TableName03);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                dr = _eFunctions.GetQueryResult(Queries.GetLostProdCodeListQuery(_eFunctions.dbReference, _eFunctions.dbLink));

                if (dr != null && !dr.IsClosed && dr.HasRows)
                {
                    var i = 5;
                    while (dr.Read())
                    {
                        _cells.GetCell(1, i).Value = dr["CODE"].ToString().Trim();
                        _cells.GetCell(2, i).Value = dr["DESCRIPTION"].ToString().Trim();
                        i++;
                    }
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 4 - CollectionSheet
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(4).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName04;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "COLECCIONES GENERADAS - ELLIPSE 8";
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


                //GENERAL

                _cells.GetCell(01, TitleRow01).Value = "EQUIP_NO";
                _cells.GetCell(02, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(02, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(03, TitleRow01).Value = "COMP_MOD_CODE";
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(04, TitleRow01).Value = "START_DATE";
                _cells.GetCell(04, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(05, TitleRow01).Value = "START_TIME";
                _cells.GetCell(05, TitleRow01).AddComment("hhmm");
                _cells.GetCell(06, TitleRow01).Value = "FINISH_DATE";
                _cells.GetCell(06, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(06, TitleRow01).AddComment("yyyyMMdd - Requerido si se usa COLLECTION");
                _cells.GetCell(07, TitleRow01).Value = "FINISH_TIME";
                _cells.GetCell(07, TitleRow01).AddComment("hhmm");
                _cells.GetCell(08, TitleRow01).Value = "ELAPSED";
                _cells.GetCell(08, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(09, TitleRow01).Value = "COLLECTION";
                _cells.GetCell(09, TitleRow01).AddComment(ShiftConstants.ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftConstants.ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftConstants.ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftConstants.ShiftCodes.DailyMorningCode + ": Dia 06-06");
                _cells.GetCell(09, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow01).Value = "SHIFT";
                _cells.GetCell(10, TitleRow01).AddComment("Este campo será ignorado si usa alguna colección");
                _cells.GetCell(11, TitleRow01).Value = "EVENT_TYPE";
                _cells.GetCell(11, TitleRow01).AddComment("LOST/DOWN");
                _cells.GetCell(12, TitleRow01).Value = "EVENT_CODE";
                _cells.GetCell(13, TitleRow01).Value = "EVENT_DESC";
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(14, TitleRow01).Value = "WORKORDER/COMENTARIO";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(14, TitleRow01).AddComment("WorkOrder para Down, Comentario para Lost");


                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01-1, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01-1, TitleRow01 + 1), TableName04);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();


                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        private void FormatSheetPbv()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(SheetName04);
                _cells.CreateNewWorksheet(ValidationSheetName);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameP01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "DOWN TIME & LOST PRODUCTION PBV - ELLIPSE 8";
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
                //Cells.GetCell("K5").Value = "REQUERIDO ADICIONAL";
                //Cells.GetCell("K5").Style = Cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetCell("A3").Value = "DISTRICT";

                _cells.SetValidationList(_cells.GetCell("B3"), DistrictConstants.GetDistrictList(), ValidationSheetName, 1);
                _cells.GetCell("B3").Value = "ICOR";

                var equipTypeList = new List<string> {"EQUIPMENT", "EGI", "LIST TYPE", "PROD.UNIT"};

                _cells.SetValidationList(_cells.GetCell("A4"), equipTypeList, ValidationSheetName, 2);
                _cells.GetCell("A4").Value = "EQUIPMENT";

                _cells.GetCell("A5").Value = "LIST ID";
                _cells.GetCell("A5").AddComment("Solo para List Type");

                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("C3").Value = "DESDE";
                _cells.GetCell("C4").Value = "HASTA";
                _cells.GetRange("C3", "C5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D5").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("D3").AddComment("yyyyMMdd");
                _cells.GetCell("D4").AddComment("yyyyMMdd");

                _cells.GetCell("C5").Value = "TYPE";
                var dataTypeList = new List<string> {"DOWN", "LOST", "DOWN & LOST"};

                _cells.SetValidationList(_cells.GetCell("D5"), dataTypeList, ValidationSheetName, 3);
                _cells.GetCell("D5").Value = "DOWN & LOST";
                //Valores predeterminados
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + "0101";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);


                _cells.GetRange(1, TitleRow01, ResultColumnP01, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetRange("A3", "D5").NumberFormat = NumberFormatConstants.Text;

                //GENERAL

                _cells.GetCell(01, TitleRow01).Value = "EQUIP_NO";
                _cells.GetCell(02, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(02, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(03, TitleRow01).Value = "COMP_MOD_CODE";
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(04, TitleRow01).Value = "START_DATE";
                _cells.GetCell(04, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(05, TitleRow01).Value = "START_TIME";
                _cells.GetCell(05, TitleRow01).AddComment("hhmm");
                _cells.GetCell(06, TitleRow01).Value = "FINISH_DATE";
                _cells.GetCell(06, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(06, TitleRow01).AddComment("yyyyMMdd - Requerido si se usa COLLECTION");
                _cells.GetCell(07, TitleRow01).Value = "FINISH_TIME";
                _cells.GetCell(07, TitleRow01).AddComment("hhmm");
                _cells.GetCell(08, TitleRow01).Value = "ELAPSED";
                _cells.GetCell(08, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(09, TitleRow01).Value = "COLLECTION";
                _cells.GetCell(09, TitleRow01).AddComment(ShiftConstants.ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftConstants.ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftConstants.ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftConstants.ShiftCodes.DailyMorningCode + ": Dia 06-06");
                _cells.GetCell(09, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow01).Value = "SHIFT";
                _cells.GetCell(10, TitleRow01).AddComment("Este campo será ignorado si usa alguna colección");
                _cells.GetCell(11, TitleRow01).Value = "EVENT_TYPE";
                _cells.GetCell(11, TitleRow01).AddComment("LOST/DOWN");
                _cells.GetCell(12, TitleRow01).Value = "EVENT_CODE";
                _cells.GetCell(13, TitleRow01).Value = "EVENT_DESC";
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(14, TitleRow01).Value = "COMENTARIO";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(15, TitleRow01).Value = "EVENTO";
                _cells.GetCell(15, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(16, TitleRow01).Value = "SYMPTOMID";
                _cells.GetCell(16, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(17, TitleRow01).Value = "ASSETTYPEID";
                _cells.GetCell(17, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(18, TitleRow01).Value = "STATUSCHANGEID";
                _cells.GetCell(18, TitleRow01).Style = StyleConstants.TitleOptional;

                _cells.GetCell(ResultColumnP01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumnP01, TitleRow01).Style = StyleConstants.TitleResult;

                //Adición de validaciones de campo
                var collectionList = new List<string>
                {
                    ShiftConstants.ShiftCodes.HourToHourCode,
                    ShiftConstants.ShiftCodes.DailyZeroCode,
                    ShiftConstants.ShiftCodes.DailyMorningCode,
                    ShiftConstants.ShiftCodes.DayNightCode
                };
                _cells.SetValidationList(_cells.GetRange(09, TitleRow01 + 1, 09, TitleRow01 + 101), collectionList, ValidationSheetName, 4);

                var typeEvent = new List<string> {"DOWN", "LOST"};
                _cells.SetValidationList(_cells.GetRange(11, TitleRow01 + 1, 11, TitleRow01 + 101), typeEvent, ValidationSheetName, 5);


                _cells.GetRange(1, TitleRow01 + 1, ResultColumnP01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumnP01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 2
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);

                _cells.GetCell("B1").Value = "DOWN TIME CODE LIST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "I2");


                _cells.GetCell("A4").Value = "CÓDIGO";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell("B4").Value = "DESCRIPCIÓN";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetRange("A4", "B4").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.FormatAsTable(_cells.GetRange("A4", "B5"), TableName02);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                var dr = _eFunctions.GetQueryResult(Queries.GetDownTimeCodeListQuery(_eFunctions.dbReference, _eFunctions.dbLink));

                if (dr != null && !dr.IsClosed && dr.HasRows)
                {
                    var i = 5;
                    while (dr.Read())
                    {
                        _cells.GetCell(1, i).Value = "'" + dr["CODE"].ToString().Trim();
                        _cells.GetCell(2, i).Value = dr["DESCRIPTION"].ToString().Trim();
                        i++;
                    }
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();



                //CONSTRUYO LA HOJA 3
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);

                _cells.GetCell("B1").Value = "LOST PRODUCTION CODE LIST - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "I2");


                _cells.GetCell("A4").Value = "CÓDIGO";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell("B4").Value = "DESCRIPCIÓN";
                _cells.GetCell("A5").NumberFormat = NumberFormatConstants.Text;
                _cells.GetRange("A4", "B4").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.FormatAsTable(_cells.GetRange("A4", "B5"), TableName03);

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
                dr = _eFunctions.GetQueryResult(Queries.GetLostProdCodeListQuery(_eFunctions.dbReference, _eFunctions.dbLink));

                if (dr != null && !dr.IsClosed && dr.HasRows)
                {
                    var i = 5;
                    while (dr.Read())
                    {
                        _cells.GetCell(1, i).Value = dr["CODE"].ToString().Trim();
                        _cells.GetCell(2, i).Value = dr["DESCRIPTION"].ToString().Trim();
                        i++;
                    }
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 4 - CollectionSheet
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(4).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName04;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "COLECCIONES GENERADAS - ELLIPSE 8";
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


                //GENERAL

                _cells.GetCell(01, TitleRow01).Value = "EQUIP_NO";
                _cells.GetCell(02, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(02, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(03, TitleRow01).Value = "COMP_MOD_CODE";
                _cells.GetCell(03, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(04, TitleRow01).Value = "START_DATE";
                _cells.GetCell(04, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(05, TitleRow01).Value = "START_TIME";
                _cells.GetCell(05, TitleRow01).AddComment("hhmm");
                _cells.GetCell(06, TitleRow01).Value = "FINISH_DATE";
                _cells.GetCell(06, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(06, TitleRow01).AddComment("yyyyMMdd - Requerido si se usa COLLECTION");
                _cells.GetCell(07, TitleRow01).Value = "FINISH_TIME";
                _cells.GetCell(07, TitleRow01).AddComment("hhmm");
                _cells.GetCell(08, TitleRow01).Value = "ELAPSED";
                _cells.GetCell(08, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(09, TitleRow01).Value = "COLLECTION";
                _cells.GetCell(09, TitleRow01).AddComment(ShiftConstants.ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftConstants.ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftConstants.ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftConstants.ShiftCodes.DailyMorningCode + ": Dia 06-06");
                _cells.GetCell(09, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow01).Value = "SHIFT";
                _cells.GetCell(10, TitleRow01).AddComment("Este campo será ignorado si usa alguna colección");
                _cells.GetCell(11, TitleRow01).Value = "EVENT_TYPE";
                _cells.GetCell(11, TitleRow01).AddComment("LOST/DOWN");
                _cells.GetCell(12, TitleRow01).Value = "EVENT_CODE";
                _cells.GetCell(13, TitleRow01).Value = "EVENT_DESC";
                _cells.GetCell(13, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(14, TitleRow01).Value = "COMENTARIO";
                _cells.GetCell(14, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(15, TitleRow01).Value = "EVENTO";
                _cells.GetCell(15, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(16, TitleRow01).Value = "SYMPTOMID";
                _cells.GetCell(16, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(17, TitleRow01).Value = "ASSETTYPEID";
                _cells.GetCell(17, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(18, TitleRow01).Value = "STATUSCHANGEID";
                _cells.GetCell(18, TitleRow01).Style = StyleConstants.TitleOptional;


                _cells.GetRange(1, TitleRow01 + 1, ResultColumnP01 - 1, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumnP01 - 1, TitleRow01 + 1), TableName04);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void ReviewDownLostPbv()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var cells = new ExcelStyleCells(_excelApp, false);
                _cells.SetCursorWait();
                cells.ClearTableRange(TableName01);
                
                var startDate = _cells.GetNullIfTrimmedEmpty(cells.GetCell("D3").Value);
                var endDate = _cells.GetNullIfTrimmedEmpty(cells.GetCell("D4").Value);

                _eFunctions.SetDBSettings(EnviromentConstants.ScadaRdb);
                var sqlQuery = Queries.GetDownLostPbv(startDate, endDate);

                var reader = _eFunctions.GetSqlQueryResult(sqlQuery);
                var currentRow = TitleRow01 + 1;

                if (reader == null || reader.IsClosed || !reader.HasRows)
                    return;
                while (reader.Read())
                {

                    cells.GetRange(01, currentRow, 18, currentRow).Style = StyleConstants.Normal;
                    cells.GetCell(01, currentRow).Value = "'" + reader[0];
                    cells.GetCell(02, currentRow).Value = "'" + reader[1];
                    cells.GetCell(03, currentRow).Value = "'" + reader[2];
                    cells.GetCell(04, currentRow).Value = "'" + reader[3];
                    cells.GetCell(05, currentRow).Value = "'" + reader[4];
                    cells.GetCell(06, currentRow).Value = "'" + reader[5];
                    cells.GetCell(07, currentRow).Value = "'" + reader[6];
                    cells.GetCell(08, currentRow).Value = "'" + reader[7];
                    cells.GetCell(09, currentRow).Value = "'" + reader[8];
                    cells.GetCell(10, currentRow).Value = "'" + reader[9];
                    cells.GetCell(11, currentRow).Value = "'" + reader[10];
                    cells.GetCell(12, currentRow).Value = "'" + reader[11];
                    cells.GetCell(13, currentRow).Value = "'" + reader[12];
                    cells.GetCell(14, currentRow).Value = "'" + reader[13];
                    cells.GetCell(15, currentRow).Value = "'" + reader[14];
                    cells.GetCell(16, currentRow).Value = "'" + reader[15];
                    cells.GetCell(17, currentRow).Value = "'" + reader[16];
                    cells.GetCell(18, currentRow).Value = "'" + reader[17];

                    if (string.IsNullOrWhiteSpace("" + reader[10]))
                        cells.GetCell(11, currentRow).Style = StyleConstants.Warning;
                    if (string.IsNullOrWhiteSpace("" + reader[12]))
                        cells.GetCell(13, currentRow).Style = StyleConstants.Warning;
                    if (string.IsNullOrWhiteSpace("" + reader[13]))
                        cells.GetCell(14, currentRow).Style = StyleConstants.Warning;
                    currentRow ++;
                }
                _eFunctions.CloseConnection();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLostPbv()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar obtener datos de la opción seleccionado");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        /// <summary>
        /// Consulta los registros Down y Lost del equipo especificado en la hoja
        /// </summary>
        public void ReviewDownLost()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var cells = new ExcelStyleCells(_excelApp, false);
                _cells.SetCursorWait();
                cells.ClearTableRange(TableName01);
                


                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                var districtCode = cells.GetNullIfTrimmedEmpty(cells.GetCell("B3").Value);
                var equipType = cells.GetNullIfTrimmedEmpty(cells.GetCell("A4").Value);
                var equipRef = cells.GetNullIfTrimmedEmpty(cells.GetCell("B4").Value); //Equipment, EGI, PU, ListType
                var listId = cells.GetNullIfTrimmedEmpty(cells.GetCell("B5").Value);
                var startDate = cells.GetNullIfTrimmedEmpty(cells.GetCell("D3").Value);
                var endDate = cells.GetNullIfTrimmedEmpty(cells.GetCell("D4").Value);
                var dataType = cells.GetNullIfTrimmedEmpty(cells.GetCell("D5").Value);

                List<string> equipmentList;
                if (equipType.Equals("EQUIPMENT"))
                    equipmentList = EquipmentActions.GetEquipmentList(_eFunctions, districtCode, equipRef);
                else if (equipType.Equals("EGI"))
                    equipmentList = EquipmentActions.GetEgiEquipments(_eFunctions, equipRef);
                else if (equipType.Equals("LIST TYPE"))
                    equipmentList = EquipmentActions.GetListEquipments(_eFunctions, equipRef, listId);
                else if (equipType.Equals("PROD.UNIT"))
                    equipmentList = EquipmentActions.GetProductiveUnitEquipments(_eFunctions, districtCode, equipRef);
                else
                    equipmentList = new List<string>();

                var i = TitleRow01 + 1;

                foreach (var equip in equipmentList)
                {
                    if (dataType.Equals("DOWN") || dataType.Equals("DOWN & LOST"))
                    {
                        var sqlDownQuery = Queries.GetEquipmentDownQuery(_eFunctions.dbReference, _eFunctions.dbLink, equip,
                            startDate, endDate);
                        var ddr = _eFunctions.GetQueryResult(sqlDownQuery);

                        if (ddr != null && !ddr.IsClosed && ddr.HasRows)
                        {
                            while (ddr.Read())
                            {
                                cells.GetCell(01, i).Value = "'" + cells.GetEmptyIfNull(ddr["EQUIP_NO"].ToString());
                                cells.GetCell(02, i).Value = "'" + cells.GetEmptyIfNull(ddr["COMP_CODE"].ToString());
                                cells.GetCell(03, i).Value = "'" + cells.GetEmptyIfNull(ddr["COMP_MOD_CODE"].ToString());
                                cells.GetCell(04, i).Value = "'" + cells.GetEmptyIfNull(ddr["START_DATE"].ToString());
                                cells.GetCell(05, i).Value = "'" + cells.GetEmptyIfNull(ddr["START_TIME"].ToString());
                                cells.GetCell(06, i).Value = "'" + cells.GetEmptyIfNull(ddr["FINISH_DATE"].ToString());
                                cells.GetCell(07, i).Value = "'" + cells.GetEmptyIfNull(ddr["FINISH_TIME"].ToString());
                                cells.GetCell(08, i).Value = "'" + cells.GetEmptyIfNull(ddr["ELAPSED_HOURS"].ToString());
                                cells.GetCell(09, i).Value = ""; //COLLECTION TYPE
                                cells.GetCell(10, i).Value = "'" + cells.GetEmptyIfNull(ddr["SHIFT"].ToString());
                                cells.GetCell(11, i).Value = "'" + cells.GetEmptyIfNull(ddr["EVENT_TYPE"].ToString());
                                cells.GetCell(12, i).Value = "'" + cells.GetEmptyIfNull(ddr["EVENT_CODE"].ToString());
                                cells.GetCell(13, i).Value = "'" + cells.GetEmptyIfNull(ddr["DESCRIPTION"].ToString());
                                cells.GetCell(14, i).Value = "'" + cells.GetEmptyIfNull(ddr["WO_COMMENT"].ToString());
                                cells.GetCell(01, i).Select();
                                i++;
                            }
                        }
                    }
                    if (dataType.Equals("LOST") || dataType.Equals("DOWN & LOST"))
                    {
                        var lostSqlQuery = Queries.GetEquipmentLostQuery(_eFunctions.dbReference,
                            _eFunctions.dbLink, equip, startDate, endDate);
                        var ddr =
                            _eFunctions.GetQueryResult(lostSqlQuery);

                        if (ddr != null && !ddr.IsClosed && ddr.HasRows)
                        {
                            while (ddr.Read())
                            {
                                cells.GetCell(01, i).Value = "'" + cells.GetEmptyIfNull(ddr["EQUIP_NO"].ToString());
                                cells.GetCell(02, i).Value = "'" + cells.GetEmptyIfNull(ddr["COMP_CODE"].ToString());
                                cells.GetCell(03, i).Value = "'" + cells.GetEmptyIfNull(ddr["COMP_MOD_CODE"].ToString());
                                cells.GetCell(04, i).Value = "'" + cells.GetEmptyIfNull(ddr["START_DATE"].ToString());
                                cells.GetCell(05, i).Value = "'" + cells.GetEmptyIfNull(ddr["START_TIME"].ToString());
                                cells.GetCell(06, i).Value = "'" + cells.GetEmptyIfNull(ddr["FINISH_DATE"].ToString());
                                cells.GetCell(07, i).Value = "'" + cells.GetEmptyIfNull(ddr["FINISH_TIME"].ToString());
                                cells.GetCell(08, i).Value = "'" + cells.GetEmptyIfNull(ddr["ELAPSED_HOURS"].ToString());
                                cells.GetCell(09, i).Value = ""; //COLLECTION TYPE
                                cells.GetCell(10, i).Value = "'" + cells.GetEmptyIfNull(ddr["SHIFT"].ToString());
                                cells.GetCell(11, i).Value = "'" + cells.GetEmptyIfNull(ddr["EVENT_TYPE"].ToString());
                                cells.GetCell(12, i).Value = "'" + cells.GetEmptyIfNull(ddr["EVENT_CODE"].ToString());
                                cells.GetCell(13, i).Value = "'" + cells.GetEmptyIfNull(ddr["DESCRIPTION"].ToString());
                                cells.GetCell(14, i).Value = "'" + cells.GetEmptyIfNull(ddr["WO_COMMENT"].ToString());
                                cells.GetCell(01, i).Select();
                                i++;
                            }
                        }
                    }
                    _eFunctions.CloseConnection();
                }

                _eFunctions.CloseConnection();
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar obtener datos de la opción seleccionado");
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        /// <summary>
        /// Inicia la acción para la creación de los registros Down/Lost en la hoja establecida. Invoca los métodos de creación individual de Lost y Down
        /// </summary>
        public void CreateDownLost()
        {
            try
            {
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01 &
                    _excelApp.ActiveWorkbook.ActiveSheet.Name != SheetNameP01)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                
                if (drpEnviroment.SelectedItem.Label != null && !drpEnviroment.SelectedItem.Label.Equals(""))
                {
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    var cells = new ExcelStyleCells(_excelApp, false);
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                    var cellCollection = new ExcelStyleCells(_excelApp, SheetName04);

                    var i = TitleRow01 + 1;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    var resultColumn01 = (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                        ? ResultColumnP01
                        : ResultColumn01;

                    while ("" + cells.GetCell(1, i).Value != "")
                    {
                        try
                        {
                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                maxInstancesSpecified = true,
                                returnWarnings = Debugger.DebugWarnings,
                                returnWarningsSpecified = true
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService

                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            var equipNo = cells.GetEmptyIfNull(cells.GetCell(1, i).Value);
                            var compCode = cells.GetEmptyIfNull(cells.GetCell(2, i).Value);
                            var compModCode = cells.GetEmptyIfNull(cells.GetCell(3, i).Value);
                            var startDate = cells.GetEmptyIfNull(cells.GetCell(4, i).Value);
                            var startTime = cells.GetEmptyIfNull(cells.GetCell(5, i).Value);
                            var finishDate = cells.GetEmptyIfNull(cells.GetCell(6, i).Value);
                            var finishTime = cells.GetEmptyIfNull(cells.GetCell(7, i).Value);
                            var elapsed = cells.GetEmptyIfNull(cells.GetCell(8, i).Value);
                            var collection = cells.GetEmptyIfNull(cells.GetCell(9, i).Value); //collection
                            var shiftCode = cells.GetEmptyIfNull(cells.GetCell(10, i).Value);
                            var eventType = cells.GetEmptyIfNull(cells.GetCell(11, i).Value);
                            var eventCode = cells.GetEmptyIfNull(cells.GetCell(12, i).Value);
                            var eventDescription = cells.GetEmptyIfNull(cells.GetCell(13, i).Value); //solo para consulta
                            var woComment = cells.GetEmptyIfNull(cells.GetCell(14, i).Value);

                            string woEvent = null;
                            string symptomId = null;
                            string assetTypeId = null;
                            string statusChangeid = null;

                            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                            {
                                woEvent = cells.GetEmptyIfNull(cells.GetCell(15, i).Value);
                                symptomId = cells.GetEmptyIfNull(cells.GetCell(16, i).Value);
                                assetTypeId = cells.GetEmptyIfNull(cells.GetCell(17, i).Value);
                                statusChangeid = cells.GetEmptyIfNull(cells.GetCell(18, i).Value);
                            }

                            startTime = string.IsNullOrWhiteSpace(startTime) ? startTime : startTime.PadLeft(4, '0');
                            finishTime = string.IsNullOrWhiteSpace(finishTime) ? finishTime : finishTime.PadLeft(4, '0');

                            LostDownObject[] ldObject;
                            if (collection == ShiftConstants.ShiftCodes.HourToHourCode ||
                                //hora a hora (Ej. 00-01, 01-02, ..., 22-23, 23-24
                                collection == ShiftConstants.ShiftCodes.DailyZeroCode || //diaria de 00-24
                                collection == ShiftConstants.ShiftCodes.DailyMorningCode || //diaria de 06-06
                                collection == ShiftConstants.ShiftCodes.DayNightCode) //dia 06-18 y noche de 18-06
                            {
                                //si es generado por colecction
                                var startEvent = new DateTime(
                                    Convert.ToInt32(startDate.Substring(0, 4)),
                                    Convert.ToInt32(startDate.Substring(4, 2)),
                                    Convert.ToInt32(startDate.Substring(6, 2)),
                                    Convert.ToInt32(Convert.ToInt32(startTime).ToString("0000").Substring(0, 2)),
                                    Convert.ToInt32(Convert.ToInt32(startTime).ToString("0000").Substring(2, 2)),
                                    00);
                                var endEvent = new DateTime(
                                    Convert.ToInt32(finishDate.Substring(0, 4)),
                                    Convert.ToInt32(finishDate.Substring(4, 2)),
                                    Convert.ToInt32(finishDate.Substring(6, 2)),
                                    Convert.ToInt32(Convert.ToInt32(finishTime).ToString("0000").Substring(0, 2)),
                                    Convert.ToInt32(Convert.ToInt32(finishTime).ToString("0000").Substring(2, 2)),
                                    00);
                                var shiftArray =
                                    TimeOperations.GetSlots(GetTurnShifts(collection), startEvent, endEvent).ToArray();

                                ldObject = new LostDownObject[shiftArray.Length];

                                for (var j = 0; j < shiftArray.Length; j++)
                                {

                                    var dateString = TimeOperations.FormatDateToString(shiftArray[j].GetDate(),
                                        TimeOperations.DateTimeFormats.DateYYYYMMDD);
                                    var startTimeString =
                                        TimeOperations.FormatTimeToString(shiftArray[j].GetStartDateTime().TimeOfDay,
                                            TimeOperations.DateTimeFormats.TimeHHMM, "");
                                    var endTimeString =
                                        TimeOperations.FormatTimeToString(shiftArray[j].GetEndDateTime().TimeOfDay,
                                            TimeOperations.DateTimeFormats.TimeHHMM, "");

                                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                                        ldObject[j] = new LostDownObject(equipNo, compCode, compModCode, dateString,
                                            startTimeString, endTimeString,
                                            null, shiftArray[j].ShiftCode, eventCode, eventDescription, woComment,
                                            woEvent, symptomId, assetTypeId, statusChangeid);
                                    else
                                        ldObject[j] = new LostDownObject(equipNo, compCode, compModCode, dateString,
                                            startTimeString, endTimeString,
                                            null, shiftArray[j].ShiftCode, eventCode, eventDescription, woComment);

                                    AddDownLostToTableCollection(ldObject[j], eventType.ToUpper(), cellCollection);

                                }

                            }
                            else
                            {
                                ldObject = new LostDownObject[1];
                                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                                    ldObject[0] = new LostDownObject(equipNo, compCode, compModCode, startDate,
                                        startTime, finishTime, elapsed, shiftCode, eventCode, eventDescription,
                                        woComment, woEvent, symptomId, assetTypeId, statusChangeid);
                                else
                                    ldObject[0] = new LostDownObject(equipNo, compCode, compModCode, startDate,
                                        startTime, finishTime, elapsed, shiftCode, eventCode, eventDescription,
                                        woComment);
                            }
                            if ((eventType.ToUpper().Equals("DOWN") || eventType.ToUpper().Equals("D")) &&
                                ldObject != null)
                                CreateDownRegister(opSheet, proxySheet, ldObject, _ignoreDuplicate);
                            else if ((eventType.ToUpper().Equals("LOST") || eventType.ToUpper().Equals("L")) &&
                                     ldObject != null)
                                CreateLostRegister(opSheet, proxySheet, ldObject, _ignoreDuplicate);

                            cells.GetCell(resultColumn01, i).Value = "SUCCESS";
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Success;
                            cells.GetCell(resultColumn01, i).Select();
                        }
                        catch (Exception ex)
                        {
                            cells.GetCell(resultColumn01, i).Value = "ERROR: " + ex.Message;
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Error;
                            cells.GetCell(resultColumn01, i).Select();
                            Debugger.LogError("RibbonEllipse:CreateDownLost()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            i++;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:CreateDownLost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _woDownOriginator = null;//solo aplica para Downs de SCADA
            }
        }
        private void GenerateCollectionList()
        {
            try
            {
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01 &
                    _excelApp.ActiveWorkbook.ActiveSheet.Name != SheetNameP01)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                
                if (drpEnviroment.SelectedItem.Label != null && !drpEnviroment.SelectedItem.Label.Equals(""))
                {
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    var cells = new ExcelStyleCells(_excelApp, false);
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                    var cellCollection = new ExcelStyleCells(_excelApp, SheetName04);
                    var resultColumn01 = (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                        ? ResultColumnP01
                        : ResultColumn01;
                    var i = TitleRow01 + 1;

                    while ("" + cells.GetCell(1, i).Value != "")
                    {
                        try
                        {

                            var equipNo = cells.GetEmptyIfNull(cells.GetCell(1, i).Value);
                            var compCode = cells.GetEmptyIfNull(cells.GetCell(2, i).Value);
                            var compModCode = cells.GetEmptyIfNull(cells.GetCell(3, i).Value);
                            var startDate = cells.GetEmptyIfNull(cells.GetCell(4, i).Value);
                            var startTime = cells.GetEmptyIfNull(cells.GetCell(5, i).Value);
                            var finishDate = cells.GetEmptyIfNull(cells.GetCell(6, i).Value);
                            var finishTime = cells.GetEmptyIfNull(cells.GetCell(7, i).Value);
                            //var elapsed = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value);
                            var collection = cells.GetEmptyIfNull(cells.GetCell(9, i).Value); //collection
                            //var shiftCode = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value);
                            var eventType = cells.GetEmptyIfNull(cells.GetCell(11, i).Value);
                            var eventCode = cells.GetEmptyIfNull(cells.GetCell(12, i).Value);
                            var eventDescription = cells.GetEmptyIfNull(cells.GetCell(13, i).Value);
                                //solo para consulta
                            var woComment = cells.GetEmptyIfNull(cells.GetCell(14, i).Value);

                            string woEvent = null;
                            string symptomId = null;
                            string assetTypeId = null;
                            string statusChangeid = null;

                            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                            {
                                woEvent = cells.GetEmptyIfNull(cells.GetCell(15, i).Value);
                                symptomId = cells.GetEmptyIfNull(cells.GetCell(16, i).Value);
                                assetTypeId = cells.GetEmptyIfNull(cells.GetCell(17, i).Value);
                                statusChangeid = cells.GetEmptyIfNull(cells.GetCell(18, i).Value);
                            }


                            if (collection == ShiftConstants.ShiftCodes.HourToHourCode ||
                                //hora a hora (Ej. 00-01, 01-02, ..., 22-23, 23-24
                                collection == ShiftConstants.ShiftCodes.DailyZeroCode || //diaria de 00-24
                                collection == ShiftConstants.ShiftCodes.DailyMorningCode || //diaria de 06-06
                                collection == ShiftConstants.ShiftCodes.DayNightCode) //dia 06-18 y noche de 18-06
                            {
                                //si es generado por colecction
                                var startEvent = new DateTime(
                                    Convert.ToInt32(startDate.Substring(0, 4)),
                                    Convert.ToInt32(startDate.Substring(4, 2)),
                                    Convert.ToInt32(startDate.Substring(6, 2)),
                                    Convert.ToInt32(Convert.ToInt32(startTime).ToString("0000").Substring(0, 2)),
                                    Convert.ToInt32(Convert.ToInt32(startTime).ToString("0000").Substring(2, 2)),
                                    00);
                                var endEvent = new DateTime(
                                    Convert.ToInt32(finishDate.Substring(0, 4)),
                                    Convert.ToInt32(finishDate.Substring(4, 2)),
                                    Convert.ToInt32(finishDate.Substring(6, 2)),
                                    Convert.ToInt32(Convert.ToInt32(finishTime).ToString("0000").Substring(0, 2)),
                                    Convert.ToInt32(Convert.ToInt32(finishTime).ToString("0000").Substring(2, 2)),
                                    00);
                                var shiftArray =
                                    TimeOperations.GetSlots(GetTurnShifts(collection), startEvent, endEvent).ToArray();

                                var ldObject = new LostDownObject[shiftArray.Length];

                                for (var j = 0; j < shiftArray.Length; j++)
                                {

                                    var dateString = TimeOperations.FormatDateToString(shiftArray[j].GetDate(),
                                        TimeOperations.DateTimeFormats.DateYYYYMMDD);
                                    var startTimeString =
                                        TimeOperations.FormatTimeToString(shiftArray[j].GetStartDateTime().TimeOfDay,
                                            TimeOperations.DateTimeFormats.TimeHHMM, "");
                                    var endTimeString =
                                        TimeOperations.FormatTimeToString(shiftArray[j].GetEndDateTime().TimeOfDay,
                                            TimeOperations.DateTimeFormats.TimeHHMM, "");

                                    if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                                        ldObject[j] = new LostDownObject(equipNo, compCode, compModCode, dateString,
                                            startTimeString, endTimeString,
                                            null, shiftArray[j].ShiftCode, eventCode, eventDescription, woComment,
                                            woEvent, symptomId, assetTypeId, statusChangeid);
                                    else
                                        ldObject[j] = new LostDownObject(equipNo, compCode, compModCode, dateString,
                                            startTimeString, endTimeString,
                                            null, shiftArray[j].ShiftCode, eventCode, eventDescription, woComment);

                                    AddDownLostToTableCollection(ldObject[j], eventType.ToUpper(), cellCollection);

                                }

                            }

                            cells.GetCell(resultColumn01, i).Value = "COLECCIÓN";
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Success;
                            cells.GetCell(resultColumn01, i).Select();
                        }
                        catch (Exception ex)
                        {
                            cells.GetCell(resultColumn01, i).Value = "ERROR: " + ex.Message;
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Error;
                            cells.GetCell(resultColumn01, i).Select();
                            Debugger.LogError("RibbonEllipse:CreateDownLost()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            i++;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:CreateDownLost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        private void AddDownLostToTableCollection(LostDownObject lostDownObject, string eventType, ExcelStyleCells cells)
        {
            //Escribo el objeto de colecction en la tabla collection
            var tableRange = cells.GetRange(TableName04);
            var row = tableRange.ListObject.ListColumns[1].Range.Row + tableRange.ListObject.ListRows.Count + 1;
            cells.GetCell(01, row).Value = "'" + lostDownObject.EquipNo;
            cells.GetCell(02, row).Value = "'" + lostDownObject.CompCode;
            cells.GetCell(03, row).Value = "'" + lostDownObject.CompModCode;
            cells.GetCell(04, row).Value = "'" + lostDownObject.Date;
            cells.GetCell(05, row).Value = "'" + lostDownObject.StartTime;
            cells.GetCell(06, row).Value = "'";
            cells.GetCell(07, row).Value = "'" + lostDownObject.FinishTime;
            cells.GetCell(08, row).Value = "'" + lostDownObject.Elapsed;
            cells.GetCell(09, row).Value = "'";
            cells.GetCell(10, row).Value = "'" + lostDownObject.ShiftCode;
            cells.GetCell(11, row).Value = "'" + eventType;
            cells.GetCell(12, row).Value = "'" + lostDownObject.EventCode;
            cells.GetCell(13, row).Value = "'" + lostDownObject.EventDescription;
            cells.GetCell(14, row).Value = "'" + lostDownObject.WoComment;
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetNameP01) return;
            cells.GetCell(15, row).Value = "'" + lostDownObject.WoEvent;
            cells.GetCell(16, row).Value = "'" + lostDownObject.SymptomId;
            cells.GetCell(17, row).Value = "'" + lostDownObject.AssetTypeId;
            cells.GetCell(18, row).Value = "'" + lostDownObject.StatusChangeid;
        }

        /// <summary>
        /// Crea un registro Down en el MSO420 para la colección de objetos down del turno establecido
        /// </summary>
        /// <param name="opContext">Screen.OperationContext: Contexto de operación del Screen Service</param>
        /// <param name="proxySheet">Screen.ScreenService: Servicio de Screen Service a utilizar</param>
        /// <param name="ldObject">LostDownObject[] : Arreglo de objetos a adicionar para Down</param>
        /// <param name="ignoreDuplicate">bool: true para ignorar duplicado para el cargue de colección</param>
        public void CreateDownRegister(Screen.OperationContext opContext, Screen.ScreenService proxySheet, LostDownObject[] ldObject, bool ignoreDuplicate = false)
        {
            foreach (var down in ldObject)
            {
                try
                {
                    proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
                    _eFunctions.RevertOperation(opContext, proxySheet);
                    //ejecutamos el programa
                    var reply = proxySheet.executeScreen(opContext, "MSO420");
                    //Validamos el ingreso
                    if (reply.mapName != "MSM420A") continue;

                    //se adicionan los valores a los campos
                    var arrayFields = new ArrayScreenNameValue();
                    arrayFields.Add("PLANT_NO1I", down.EquipNo);
                    arrayFields.Add("STAT_DATE1I", down.Date);
                    arrayFields.Add("SHIFT1I", down.ShiftCode);

                    var request = new Screen.ScreenSubmitRequestDTO
                    {
                        screenFields = arrayFields.ToArray(),
                        screenKey = "1"
                    };
                    reply = proxySheet.submit(opContext, request);

                    //no hay errores ni advertencias
                    if (reply != null && !_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply))
                    {
                        var replyFields = new ArrayScreenNameValue(reply.screenFields);
                        var k = 1;
                        //hasta ubicar un campo vacío o el mismo registro a cargar
                        while (_cells.GetEmptyIfNull(replyFields.GetField("DOWN_TIME_CODE1I" + k).value) != ""
                            & (_cells.GetEmptyIfNull(replyFields.GetField("DOWN_TIME_CODE1I" + k).value) + _cells.GetEmptyIfNull(replyFields.GetField("STOP_TIME1I" + k).value) != down.EventCode + down.StartTime.Substring(0, 2) + ":" + down.StartTime.Substring(2, 2)))
                        {
                            k++;

                            if (k <= 10) continue;
                            k = 1;
                            //envíe a la siguiente pantalla
                            request = new Screen.ScreenSubmitRequestDTO {screenKey = "1"};
                            proxySheet.submit(opContext, request);
                            replyFields = new ArrayScreenNameValue(reply.screenFields);
                        }
                        
                        if (down.WoEvent != null && WorkOrderActions.FetchWorkOrder(_eFunctions, "", down.WoEvent) == null)
                        {
                            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                            var wo = new WorkOrder();
                            var opContextWo = new OperationContext
                            {
                                district = opContext.district,
                                position = opContext.position,
                                maxInstances = opContext.maxInstances,
                                maxInstancesSpecified = true,
                                returnWarnings = opContext.returnWarnings,
                                returnWarningsSpecified = true
                            };


                            wo.workGroup = "PTOEVNT";
                            var workNo = down.WoEvent;
                            if (workNo != null)
                            {
                                if (workNo.Length == 2) //prefijo
                                    wo.SetWorkOrderDto(workNo, null);
                                else
                                    wo.SetWorkOrderDto(workNo);
                            }

                            wo.workOrderDesc = down.EventDescription;
                            wo.equipmentNo = down.EquipNo;
                            wo.compCode = down.CompCode;
                            wo.compModCode = down.CompModCode;
                            wo.workOrderType = "RE";
                            wo.maintenanceType = "CO";

                            //DETAILS
                            if (string.IsNullOrWhiteSpace(_woDownOriginator))
                                _woDownOriginator = InputBox.GetValue("Work Order", "Ingrese Originator Id:", "USERNAME");
                            wo.originatorId = _woDownOriginator;
                            wo.origPriority = "P3";
                            //PLANNING
                            wo.planPriority = "P3";
                            wo.requiredByDate = down.Date;
                            wo.requiredByTime = down.StartTime;
                            //COST

                            //JOB_CODES
                            wo.jobCode1 = down.SymptomId;
                            wo.jobCode2 = down.AssetTypeId;
                            wo.jobCode3 = down.StatusChangeid;

                            WorkOrderActions.CreateWorkOrder(urlService, opContextWo, wo);
                        }

                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("DOWN_TIME_CODE1I" + k, down.EventCode);
                        arrayFields.Add("STOP_TIME1I" + k, down.StartTime); //inicio del down
                        arrayFields.Add("START_TIME1I" + k, down.FinishTime); //fin del down
                        arrayFields.Add("WORK_ORDER1I" + k, down.WoEvent ?? down.WoComment);//si el WoEvent es nulo es porque no es un evento de SCADA PBV si no un Down regular
                        arrayFields.Add("COMP_CODE1I" + k, down.CompCode);
                        arrayFields.Add("MODIFIER1I" + k, down.CompModCode);

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opContext, request);

                        while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM420A" &&
                               (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" ||
                                reply.functionKeys.StartsWith("XMIT-WARNING")))
                        {

                            request = new Screen.ScreenSubmitRequestDTO();
                            if (reply.message == "SHIFT START / STOP TIMES OVERLAP")
                            {
                                var rowToDelete = reply.currentCursorFieldName.Length == 12 ? reply.currentCursorFieldName.Substring(reply.currentCursorFieldName.Length - 1, 1) : reply.currentCursorFieldName.Substring(reply.currentCursorFieldName.Length - 2, 2);
                                arrayFields = new ArrayScreenNameValue();
                                arrayFields.Add("ACTION1I" + rowToDelete, "D");
                                request.screenFields = arrayFields.ToArray();
                            }
                            request.screenKey = "1";
                            reply = proxySheet.submit(opContext, request);


                        }
                        if (reply != null && (_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM420A"))
                            throw new ArgumentException(reply.message);


                    }
                    else if (reply != null) throw new Exception(reply.message);
                    else throw new Exception(@"No se ha podido obtener respuesta del servidor");
                }
                catch (Exception ex)
                {
                    if (!ignoreDuplicate || !ex.Message.Equals("X2:0018 - DUPLICATE ENTRY"))
                        throw new Exception(ex.Message);
                }
            }
        }

        /// <summary>
        /// Crea un registro Lost en el MSO470 para la colección de objetos lost del turno establecido
        /// </summary>
        /// <param name="opContext">Screen.OperationContext: Contexto de operación del Screen Service</param>
        /// <param name="proxySheet">Screen.ScreenService: Servicio de Screen Service a utilizar</param>
        /// <param name="ldObject">LostDownObject[] : Arreglo de objetos a adicionar para Lost</param>
        /// <param name="ignoreDuplicate"></param>
        public void CreateLostRegister(Screen.OperationContext opContext, Screen.ScreenService proxySheet, LostDownObject[] ldObject, bool ignoreDuplicate = false)
        {
            foreach (var lost in ldObject)
            {
                try
                {
                    proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
                    _eFunctions.RevertOperation(opContext, proxySheet);
                    //ejecutamos el programa
                    var reply = proxySheet.executeScreen(opContext, "MSO470");
                    //Validamos el ingreso
                    if (reply.mapName != "MSM470A") continue;

                    //se adicionan los valores a los campos
                    var arrayFields = new ArrayScreenNameValue();
                    arrayFields.Add("EQUIP_NO1I", lost.EquipNo);
                    arrayFields.Add("REV_STAT_DATE1I", lost.Date);
                    arrayFields.Add("SHIFT_CODE1I", lost.ShiftCode);

                    var request = new Screen.ScreenSubmitRequestDTO
                    {
                        screenFields = arrayFields.ToArray(),
                        screenKey = "1"
                    };
                    reply = proxySheet.submit(opContext, request);

                    //no hay errores ni advertencias
                    if (reply != null && !_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply))
                    {
                        var replyFields = new ArrayScreenNameValue(reply.screenFields);
                        var k = 1;
                        //hasta ubicar un campo vacío
                        while (_cells.GetEmptyIfNull(replyFields.GetField("LOST_PROD_CODE1I" + k).value) != ""
                            & (_cells.GetEmptyIfNull(replyFields.GetField("LOST_PROD_CODE1I" + k).value) + _cells.GetEmptyIfNull(replyFields.GetField("STOP_TIME1I" + k).value) != lost.EventCode + lost.StartTime.Substring(0, 2) + ":" + lost.StartTime.Substring(2, 2)))
                        {
                            k++;

                            if (k <= 10) continue;
                            k = 1;
                            //envíe a la siguiente pantalla
                            request = new Screen.ScreenSubmitRequestDTO {screenKey = "1"};
                            reply = proxySheet.submit(opContext, request);
                            replyFields = new ArrayScreenNameValue(reply.screenFields);
                        }

                        arrayFields = new ArrayScreenNameValue();
                        arrayFields.Add("LOST_PROD_CODE1I" + k, lost.EventCode);
                        arrayFields.Add("STOP_TIME1I" + k, lost.StartTime); //inicio del lost
                        arrayFields.Add("START_TIME1I" + k, lost.FinishTime); //fin del lost

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = proxySheet.submit(opContext, request);

                        while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM470A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
                        {
                            request = new Screen.ScreenSubmitRequestDTO();
                            if (reply.message == "SHIFT START / STOP TIMES OVERLAP")
                            {
                                var rowToDelete = reply.currentCursorFieldName.Length == 12 ? reply.currentCursorFieldName.Substring(reply.currentCursorFieldName.Length - 1, 1) : reply.currentCursorFieldName.Substring(reply.currentCursorFieldName.Length - 2, 2);
                                arrayFields = new ArrayScreenNameValue();
                                arrayFields.Add("ACTION1I" + rowToDelete, "D");
                                request.screenFields = arrayFields.ToArray();
                            }
                            request.screenKey = "1";
                            reply = proxySheet.submit(opContext, request);
                        }

                        if (reply != null && (_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM470A"))
                            throw new ArgumentException(reply.message);

                        //creación/actualización del comentario
                        if (string.IsNullOrWhiteSpace(lost.WoComment)) continue;
                        var arrayValues = new ArrayScreenNameValue(reply.screenFields);
                        lost.EquipNo = arrayValues.GetField("EQUIP_NO1I").value;
                        var dr = _eFunctions.GetQueryResult(Queries.GetSingleLostQuery(_eFunctions.dbReference,_eFunctions.dbLink, lost.EquipNo, lost.EventCode, lost.Date, lost.ShiftCode,lost.StartTime, lost.FinishTime));
                        var stdTextId = "";
                        if (dr != null && dr.HasRows)
                            while (dr.Read())
                                stdTextId = _cells.GetEmptyIfNull(dr["STD_KEY"].ToString());
                        else
                            throw new Exception(@"No se ha encontrado registro Lost para creación de comentario");
                        dr.Close();

                        var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                        stdTextId = "LP" + stdTextId;
                        var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

                        var textResult = StdText.SetText(urlService, StdText.GetCustomOpContext(district, _frmAuth.EllipsePost, 100, false), stdTextId, lost.WoComment);

                        if (!textResult)
                            throw new KeyNotFoundException(
                                "No se ha podido crear el comentario de uno de los registro Lost seleccionados");
                    }
                    else if (reply != null) throw new Exception(reply.message);
                    else throw new Exception(@"No se ha podido obtener respuesta del servidor");
                }
                catch (Exception ex)
                {
                    if (!ignoreDuplicate || !ex.Message.Equals("X2:0018 - DUPLICATE ENTRY"))
                        throw new Exception(ex.Message);
                }
            }
        }

        /// <summary>
        /// Inicia la acción para la eliminación de los registros Down/Lost en la hoja establecida. Invoca los métodos de eliminación individual de Lost y Down
        /// </summary>
        public void DeleteDownLost()
        {
            try
            {
                //si no se está en la hoja correspondiente
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01)
                    throw new InvalidOperationException("La hoja seleccionada no coincide con el modelo requerido");
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                
                _cells.SetCursorWait();
                if (drpEnviroment.SelectedItem.Label != null && !drpEnviroment.SelectedItem.Label.Equals(""))
                {
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    var cells = new ExcelStyleCells(_excelApp, false);
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                    var i = TitleRow01 + 1;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    var resultColumn01 = (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                        ? ResultColumnP01
                        : ResultColumn01;

                    while ("" + cells.GetCell(1, i).Value != "")
                    {
                        try
                        {

                            //ScreenService Opción en reemplazo de los servicios
                            var opSheet = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                maxInstancesSpecified = true,
                                returnWarnings = Debugger.DebugWarnings,
                                returnWarningsSpecified = true
                            };

                            var proxySheet = new Screen.ScreenService();
                            ////ScreenService
                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            var equipNo = cells.GetEmptyIfNull(cells.GetCell(1, i).Value);
                            var compCode = cells.GetEmptyIfNull(cells.GetCell(2, i).Value);
                            var compModCode = cells.GetEmptyIfNull(cells.GetCell(3, i).Value);
                            var startDate = cells.GetEmptyIfNull(cells.GetCell(4, i).Value);
                            var startTime = cells.GetEmptyIfNull(cells.GetCell(5, i).Value);
                            //string finishDate = Cells.GetEmptyIfNull(Cells.GetCell(6, i).Value); //collection
                            var finishTime = cells.GetEmptyIfNull(cells.GetCell(7, i).Value);
                            var elapsed = cells.GetEmptyIfNull(cells.GetCell(8, i).Value);
                            //string collection = Cells.GetEmptyIfNull(Cells.GetCell(9, i).Value);//collection
                            var shiftCode = cells.GetEmptyIfNull(cells.GetCell(10, i).Value);
                            var eventType = cells.GetEmptyIfNull(cells.GetCell(11, i).Value);
                            var eventCode = cells.GetEmptyIfNull(cells.GetCell(12, i).Value);
                            var eventDescription = cells.GetEmptyIfNull(cells.GetCell(13, i).Value); //solo consulta
                            var woComment = cells.GetEmptyIfNull(cells.GetCell(14, i).Value);
                            var woEvent = cells.GetEmptyIfNull(cells.GetCell(15, i).Value);
                            var symptomId = cells.GetEmptyIfNull(cells.GetCell(16, i).Value);
                            var assetTypeId = cells.GetEmptyIfNull(cells.GetCell(17, i).Value);
                            var statusChangeid = cells.GetEmptyIfNull(cells.GetCell(18, i).Value);


                            var ldObject = new LostDownObject(equipNo, compCode, compModCode, startDate, startTime,
                                finishTime, elapsed, shiftCode, eventCode, eventDescription, woComment, woEvent,
                                symptomId, assetTypeId, statusChangeid);
                            var resultado = false;
                            if ((eventType.ToUpper().Equals("DOWN") || eventType.ToUpper().Equals("D")) &&
                                ldObject != null)
                                resultado = DeleteDownRegister(opSheet, proxySheet, ldObject);
                            //resultado = DeleteDownRegister(opDown, ldObject); //obsoleto
                            else if ((eventType.ToUpper().Equals("LOST") || eventType.ToUpper().Equals("L")) &&
                                     ldObject != null)
                                resultado = DeleteLostRegister(opSheet, proxySheet, ldObject);
                            //resultado = DeleteDownRegister(opLost, ldObject); //obsoleto

                            if (resultado)
                            {
                                cells.GetCell(resultColumn01, i).Value = "ELIMINADO";
                                cells.GetCell(resultColumn01, i).Style = StyleConstants.Success;
                                cells.GetCell(resultColumn01, i).Select();
                            }
                            else
                            {
                                cells.GetCell(resultColumn01, i).Value = "NO SE HA REALIZADO NINGUNA ACCIÓN";
                                cells.GetCell(resultColumn01, i).Style = StyleConstants.Warning;
                                cells.GetCell(resultColumn01, i).Select();
                            }
                        }
                        catch (Exception ex)
                        {
                            cells.GetCell(resultColumn01, i).Value = "ERROR: " + ex.Message;
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Error;
                            Debugger.LogError("RibbonEllipse:DeleteDownLost()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                        }
                        finally
                        {
                            i++;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(@"Seleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:CreateDownLost()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        /// <summary>
        /// Elimina un registro Down dado
        /// </summary>
        /// <param name="opContext">Screen.OperationContext: contexto de operación del Screen Service</param>
        /// <param name="proxySheet">Screen.ScreenService: Servicio de Screen Service</param>
        /// <param name="down">LostDownObject: objeto Down para eliminación</param>
        /// <returns>true: si se realiza alguna eliminación</returns>
        public bool DeleteDownRegister(Screen.OperationContext opContext, Screen.ScreenService proxySheet, LostDownObject down)
        {
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO420");
            //Validamos el ingreso
            if (reply.mapName != "MSM420A") return false;

            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("PLANT_NO1I", down.EquipNo);
            arrayFields.Add("STAT_DATE1I", down.Date);
            arrayFields.Add("SHIFT1I", down.ShiftCode);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            //no hay errores ni advertencias
            if (reply == null || _eFunctions.CheckReplyError(reply) || _eFunctions.CheckReplyWarning(reply))
                return false;
            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            var k = 1;
            //hasta ubicar el campo que se necesita
            while (reply.mapName == "MSM420A" && _cells.GetEmptyIfNull(replyFields.GetField("DOWN_TIME_CODE1I" + k).value) != down.EventCode && down.StartTime != replyFields.GetField("STOP_TIME1I" + k).value && down.FinishTime != replyFields.GetField("START_TIME1I" + k).value)
            {
                k++;

                if (k > 10)
                {
                    k = 1;
                    //envíe a la siguiente pantalla
                    request.screenKey = "1";
                    reply = proxySheet.submit(opContext, request);
                }
            }

            if (reply.mapName != "MSM420A")
                throw new ArgumentException("No se ha encontrado el registro");

            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("ACTION1I" + k, "D");

            request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM420A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
            {
                request.screenKey = "1";
                reply = proxySheet.submit(opContext, request);
            }
            if (!_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply)) return true;
            if (reply != null) throw new ArgumentException(reply.message);
            throw new Exception(@"No se ha podido obtener respuesta del servidor");
        }
        /// <summary>
        /// Elimina un registro Lost dado
        /// </summary>
        /// <param name="opContext">Screen.OperationContext: contexto de operación del Screen Service</param>
        /// <param name="proxySheet">Screen.ScreenService: Servicio de Screen Service</param>
        /// <param name="lost">LostDownObject: objeto Lost para eliminación</param>
        /// <returns>true: si se realiza alguna eliminación</returns>
        public bool DeleteLostRegister(Screen.OperationContext opContext, Screen.ScreenService proxySheet, LostDownObject lost)
        {
            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO470");
            //Validamos el ingreso
            if (reply.mapName != "MSM470A") return false;
            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("EQUIP_NO1I", lost.EquipNo);
            arrayFields.Add("REV_STAT_DATE1I", lost.Date);
            arrayFields.Add("SHIFT_CODE1I", lost.ShiftCode);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            //no hay errores ni advertencias
            if (reply == null || _eFunctions.CheckReplyError(reply) || _eFunctions.CheckReplyWarning(reply))
                return false;
            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            var k = 1;
            //hasta ubicar el campo que se necesita
            while (reply.mapName == "MSM470A" && _cells.GetEmptyIfNull(replyFields.GetField("LOST_PROD_CODE1I" + k).value) != lost.EventCode && lost.StartTime != replyFields.GetField("STOP_TIME1I" + k).value && lost.FinishTime != replyFields.GetField("START_TIME1I" + k).value)
            {
                k++;

                if (k <= 10) continue;
                k = 1;
                //envíe a la siguiente pantalla
                request.screenKey = "1";
                reply = proxySheet.submit(opContext, request);
            }

            if (reply.mapName != "MSM470A")
                throw new ArgumentException("No se ha encontrado el registro");

            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("ACTION1I" + k, "D");

            request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            while (reply != null && !_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM470A" && (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING")))
            {
                request.screenKey = "1";
                reply = proxySheet.submit(opContext, request);
            }

            if (!_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply)) return true;
            if (reply != null) throw new ArgumentException(reply.message);
            throw new Exception(@"No se ha podido obtener respuesta del servidor");
        }

        /// <summary>
        /// Obtiene un arreglo de tipo ShiftSlot [] con la información de los turnos correspondientes a una colección de un tipo de turno (Ej: Colleción D/N tiene los turnos D de 0600 a 1800 y N de 1800 a 0600; Collección HH tiene los turnos 01 de 0000 a 0100, ..., y 24 de 2300 a 0000)
        /// </summary>
        /// <param name="shiftPeriodCode">Código de colección del Turno</param>
        /// <returns></returns>
        private static ShiftSlot[] GetTurnShifts(string shiftPeriodCode)
        {
            if (shiftPeriodCode.Equals(ShiftConstants.ShiftCodes.HourToHourCode))
                return ShiftConstants.ShiftPeriods.GetHourToHourShiftSlots();
            if (shiftPeriodCode.Equals(ShiftConstants.ShiftCodes.DailyMorningCode))
                return ShiftConstants.ShiftPeriods.GetDailyMorningSlots();
            if (shiftPeriodCode.Equals(ShiftConstants.ShiftCodes.DailyZeroCode))
                return ShiftConstants.ShiftPeriods.GetDailyZeroSlots();
            // ReSharper disable once ConvertIfStatementToReturnStatement
            if (shiftPeriodCode.Equals(ShiftConstants.ShiftCodes.DayNightCode))
                return ShiftConstants.ShiftPeriods.GetDailyNightShiftSlots();
            return null;
        }

        /// <summary>
        /// Una clase para almacenar los valores bases de un registro DOWN/LOST
        /// </summary>
        public class LostDownObject
        {
            public string EquipNo;
            public string CompCode;
            public string CompModCode;
            public string Date;
            public string StartTime;
            public string FinishTime;
            public string Elapsed;
            public string ShiftCode;
            public string EventCode;
            public string EventDescription;
            public string WoComment;
            public string WoEvent;//pbv
            public string SymptomId;//pbv
            public string AssetTypeId;//pbv
            public string StatusChangeid;//pbv

            /// <summary>
            /// Constructor de la clase
            /// </summary>
            /// <param name="equipNo">string: Número del Equipo</param>
            /// <param name="compCode">string: Código del Componente (N/A para Lost)</param>
            /// <param name="compModCode">string: Código de Modificador de componente (N/A para Lost)</param>
            /// <param name="date">string: Fecha del evento yyyyMMdd</param>
            /// <param name="startTime">string: Hora de inicio del evento hhmm</param>
            /// <param name="finishTime">string: Hora de finalización del evento hhmm</param>
            /// <param name="elapsed">string: Tiempo transcurrido (puede ser nulo)</param>
            /// <param name="shiftCode">string: código del turno</param>
            /// <param name="eventCode">string: código del evento</param>
            /// <param name="eventDescription">string: descripción del código del evento</param>
            /// <param name="woComment">string: WorkOrder para DownTime ó Texto de comentario para Lost</param>
            public LostDownObject(string equipNo, string compCode, string compModCode, string date, string startTime, string finishTime, string elapsed, string shiftCode, string eventCode, string eventDescription, string woComment)
            {
                EquipNo = equipNo;
                CompCode = compCode;
                CompModCode = compModCode;
                Date = date;
                StartTime = startTime;
                FinishTime = finishTime;
                Elapsed = elapsed;
                ShiftCode = shiftCode;
                EventCode = eventCode;
                EventDescription = eventDescription;
                WoComment = woComment;
            }

            /// <summary>
            /// Constructor de la clase con codigos de falla
            /// </summary>
            /// <param name="equipNo">string: Número del Equipo</param>
            /// <param name="compCode">string: Código del Componente (N/A para Lost)</param>
            /// <param name="compModCode">string: Código de Modificador de componente (N/A para Lost)</param>
            /// <param name="date">string: Fecha del evento yyyyMMdd</param>
            /// <param name="startTime">string: Hora de inicio del evento hhmm</param>
            /// <param name="finishTime">string: Hora de finalización del evento hhmm</param>
            /// <param name="elapsed">string: Tiempo transcurrido (puede ser nulo)</param>
            /// <param name="shiftCode">string: código del turno</param>
            /// <param name="eventCode">string: código del evento</param>
            /// <param name="eventDescription">string: descripción del código del evento</param>
            /// <param name="woComment">string: WorkOrder para DownTime ó Texto de comentario para Lost</param>
            /// <param name="woEvent">string: Orden que se genera por el evento</param>
            /// <param name="symptomId">string: codigo de falla, Sintoma</param>
            /// <param name="assetTypeId">string: codigo de falla, Componente</param>
            /// <param name="statusChangeid">string: codigo de falla, Causa</param>
            public LostDownObject(string equipNo, string compCode, string compModCode, string date, string startTime, string finishTime, string elapsed, string shiftCode, string eventCode, string eventDescription, string woComment, string woEvent, string symptomId, string assetTypeId, string statusChangeid)
            {
                EquipNo = equipNo;
                CompCode = compCode;
                CompModCode = compModCode;
                Date = date;
                StartTime = startTime;
                FinishTime = finishTime;
                Elapsed = elapsed;
                ShiftCode = shiftCode;
                EventCode = eventCode;
                EventDescription = eventDescription;
                WoComment = woComment;
                WoEvent = woEvent;
                SymptomId = symptomId;
                AssetTypeId = assetTypeId;
                StatusChangeid = statusChangeid;
            }
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread == null || !_thread.IsAlive) return;
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
    }

    public class Queries
    {
        /// <summary>
        /// Obtiene el query para el listado de Down de un equipo dado
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <param name="equipmentNo"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns>string: Query para el listado de Down de un equipo dado</returns>
        public static string GetEquipmentDownQuery(string dbreference, string dblink, string equipmentNo, string startDate, string endDate)
        {
            //Se entiende que en la tabla el STOP_TIME es la hora de inicio del Down, por esto es renombrado a START_TIME
            //y START_TIME es la hora de fin del Down por eso es renombrado a FINISH_TIME
            //Así que nuestra franja será [START_TIME, FINISH_TIME] para indiciar Inicio, Fin del Down
            var query = "" +
                " SELECT" +
                "     SUBSTR(DW.REC_EQUIP_420,2) EQUIP_NO, DW.COMP_CODE, DW.COMP_MOD_CODE," +
                "     (99999999 - DW.REV_STAT_DATE) START_DATE, DW.STOP_TIME START_TIME," +
                "     (99999999 - DW.REV_STAT_DATE) FINISH_DATE, DW.START_TIME FINISH_TIME, DW.ELAPSED_HOURS, DW.SHIFT, 'DOWN' EVENT_TYPE," +
                "     DW.DOWN_TIME_CODE EVENT_CODE, COD.TABLE_DESC DESCRIPTION, DW.WORK_ORDER WO_COMMENT," +
                "     DW.SEQUENCE_NO, DW.SHIFT_SEQ_NO" +
                "   FROM" +
                "     " + dbreference + ".MSF420" + dblink + " DW" +
                "   INNER JOIN " + dbreference + ".MSF010" + dblink + " COD ON TRIM(DW.DOWN_TIME_CODE) = TRIM(COD.TABLE_CODE)" +
                "   WHERE" +
                "     DW.REC_EQUIP_420 = 'E' ||'" + equipmentNo + "'" +
                "     AND (99999999 - DW.REV_STAT_DATE) BETWEEN '" + startDate + "' AND '" + endDate + "'" +
                "     AND COD.TABLE_TYPE = 'DT'" +
                "   ORDER BY EQUIP_NO, COMP_CODE, COMP_MOD_CODE, START_DATE, START_TIME";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }

        /// <summary>
        /// Obtiene el query para el listado de Lost Production de un equipo dado
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <param name="equipmentNo"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns>string: Query para el listado de Lost de un equipo dado</returns>
        public static string GetEquipmentLostQuery(string dbreference, string dblink, string equipmentNo, string startDate, string endDate)
        {
            //Se entiende que en la tabla el STOP_TIME es la hora de inicio del Lost, por esto es renombrado a START_TIME
            //y START_TIME es la hora de fin del Lost por eso es renombrado a FINISH_TIME
            //Así que nuestra franja será [START_TIME, FINISH_TIME] para indiciar Inicio, Fin del Lost
            var query = "" +
                " SELECT" +
                " LS.EQUIP_NO, '' COMP_CODE, '' COMP_MOD_CODE," +
                " (99999999 - LS.REV_STAT_DATE) START_DATE, LS.STOP_TIME START_TIME," +
                " (99999999 - LS.REV_STAT_DATE) FINISH_DATE, LS.START_TIME FINISH_TIME, LS.ELAPSED_HOURS, LS.SHIFT, 'LOST' EVENT_TYPE," +
                " LS.LOST_PROD_CODE EVENT_CODE, COD.TABLE_DESC DESCRIPTION, " +
                " (SELECT TRIM(LPTEXT.STD_MEDIUM_1 || LPTEXT.STD_MEDIUM_2 || LPTEXT.STD_MEDIUM_3 || LPTEXT.STD_MEDIUM_4 || LPTEXT.STD_MEDIUM_5) FROM ELLIPSE.MSF096_STD_MEDIUM LPTEXT WHERE LPTEXT.STD_TEXT_CODE = 'LP' AND ROWNUM = 1 AND LPTEXT.STD_KEY = RPAD(EQUIP_NO, 12, ' ') ||  (MOD(TO_DATE((99999999 - LS.REV_STAT_DATE), 'YYYYMMDD') - TO_DATE('19800101', 'YYYYMMDD'),9999)-1) || LS.SHIFT || LS.LOST_PROD_CODE || LS.SEQUENCE_NO) WO_COMMENT," +
                " LS.SEQUENCE_NO, LS.SHIFT_SEQ_NO" +
                "   FROM" +
                "     " + dbreference + ".MSF470" + dblink + " LS" +
                "   INNER JOIN " + dbreference + ".MSF010" + dblink + " COD ON TRIM(LS.LOST_PROD_CODE) = TRIM(COD.TABLE_CODE)" +
                "   WHERE" +
                "     LS.EQUIP_NO = '" + equipmentNo + "'" +
                "     AND (99999999 - LS.REV_STAT_DATE) BETWEEN '" + startDate + "' AND '" + endDate + "'" +
                "     AND COD.TABLE_TYPE = 'LP'" +
                "   ORDER BY EQUIP_NO, COMP_CODE, COMP_MOD_CODE, START_DATE, START_TIME";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
        /// <summary>
        /// Obtiene el query listado de Códigos Down del sistema
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <returns>string: Query para el listado de Códigos Down del Sistema</returns>
        public static string GetDownTimeCodeListQuery(string dbreference, string dblink)
        {
            var query = "" +
            " SELECT" +
            "   CASE COD.TABLE_TYPE WHEN 'LP' THEN 'LOST' WHEN 'DT' THEN 'DOWN' END EVENT_TYPE," +
            "   TRIM(COD.TABLE_CODE) CODE, " +
            "   TRIM(COD.TABLE_DESC) DESCRIPTION" +
            " FROM " + dbreference + ".MSF010" + dblink + " COD" +
            " WHERE TABLE_TYPE = 'DT'" +
            " ORDER BY TABLE_TYPE, TABLE_CODE, TABLE_DESC";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
        /// <summary>
        /// Obtiene el query listado de Códigos Lost Production del sistema
        /// </summary>
        /// <param name="dbreference">Referencia a la base de datos (Ej: MIMSPROD, ELLIPSE)</param>
        /// <param name="dblink">Link de conexión a la base de datos (Ej: @CONSULBO)</param>
        /// <returns>string: Query para el listado de Códigos Lost del Sistema</returns>
        public static string GetLostProdCodeListQuery(string dbreference, string dblink)
        {
            var query = "" +
            " SELECT" +
            "   CASE COD.TABLE_TYPE WHEN 'LP' THEN 'LOST' WHEN 'DT' THEN 'DOWN' END EVENT_TYPE," +
            "   TRIM(COD.TABLE_CODE) CODE, " +
            "   TRIM(COD.TABLE_DESC) DESCRIPTION" +
            " FROM " + dbreference + ".MSF010" + dblink + " COD" +
            " WHERE TABLE_TYPE = 'LP'" +
            " ORDER BY TABLE_TYPE, TABLE_CODE, TABLE_DESC";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
        [Obsolete("Utilizado para los métodos que utilizan los servicios de Down directamente. Se marca obsoleto porque el sistema no establece comunicación con el servicio")]
        public static string GetSingleDownQuery(string dbreference, string dblink, string equipmentNo, string downCode, string startDate, string shiftCode, string startTime, string endTime)
        {
            var query = "" +
                "   SELECT" +
                "     SUBSTR(DW.REC_EQUIP_420,2) EQUIP_NO, DW.COMP_CODE, DW.COMP_MOD_CODE," +
                "     DW.SEQUENCE_NO, DW.SHIFT_SEQ_NO, " +
                "     (99999999 - DW.REV_STAT_DATE) START_DATE, DW.STOP_TIME START_TIME," +
                "     (99999999 - DW.REV_STAT_DATE) FINISH_DATE, DW.START_TIME FINISH_TIME, DW.ELAPSED_HOURS, DW.SHIFT, 'DOWN' EVENT_TYPE," +
                "     DW.DOWN_TIME_CODE EVENT_CODE, DW.WORK_ORDER WO_COMMENT" +
                "   FROM" +
                "     " + dbreference + ".MSF420" + dblink + " DW" +
                "   WHERE" +
                "     DW.REC_EQUIP_420 = 'E' ||'" + equipmentNo + "'" +
                "     AND DW.DOWN_TIME_CODE = '" + downCode + "'" +
                "     AND (99999999 - DW.REV_STAT_DATE) = '" + startDate + "'" +
                "     AND DW.SHIFT = '" + shiftCode + "'" +
                "     AND DW.STOP_TIME = LPAD('" + startTime + "', 4, '0')" +
                "     AND DW.START_TIME = LPAD('" + endTime + "', 4, '0')";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
        public static string GetSingleLostQuery(string dbreference, string dblink, string equipmentNo, string lostCode, string startDate, string shiftCode, string startTime, string endTime)
        {
            //Se entiende que en la tabla el STOP_TIME es la hora de inicio del Lost, por esto es renombrado a START_TIME
            //y START_TIME es la hora de fin del Lost por eso es renombrado a FINISH_TIME
            //Así que nuestra franja será [START_TIME, FINISH_TIME] para indiciar Inicio, Fin del Lost
            var query = "" +
                " SELECT" +
                " LS.EQUIP_NO, '' COMP_CODE, '' COMP_MOD_CODE," +
                " LS.SEQUENCE_NO, LS.SHIFT_SEQ_NO," +
                " (99999999 - LS.REV_STAT_DATE) START_DATE, LS.STOP_TIME START_TIME," +
                " (99999999 - LS.REV_STAT_DATE) FINISH_DATE, LS.START_TIME FINISH_TIME, LS.ELAPSED_HOURS, LS.SHIFT, 'LOST' EVENT_TYPE," +
                " LS.LOST_PROD_CODE EVENT_CODE, " +
                " (SELECT TRIM(LPTEXT.STD_MEDIUM_1 || LPTEXT.STD_MEDIUM_2 || LPTEXT.STD_MEDIUM_3 || LPTEXT.STD_MEDIUM_4 || LPTEXT.STD_MEDIUM_5) FROM ELLIPSE.MSF096_STD_MEDIUM LPTEXT WHERE LPTEXT.STD_TEXT_CODE = 'LP' AND LPTEXT.STD_KEY = RPAD(EQUIP_NO, 12, ' ') ||  (MOD(TO_DATE((99999999 - LS.REV_STAT_DATE), 'YYYYMMDD') - TO_DATE('19800101', 'YYYYMMDD'),9999)-1) || LS.SHIFT || LS.LOST_PROD_CODE || LS.SEQUENCE_NO) WO_COMMENT," +
                " (RPAD(EQUIP_NO, 12, ' ') ||  (MOD(TO_DATE((99999999 - LS.REV_STAT_DATE), 'YYYYMMDD') - TO_DATE('19800101', 'YYYYMMDD'),9999)-1) || LS.SHIFT || LS.LOST_PROD_CODE || LS.SEQUENCE_NO) STD_KEY " +
                "   FROM" +
                "     " + dbreference + ".MSF470" + dblink + " LS" +
                "   INNER JOIN " + dbreference + ".MSF010" + dblink + " COD ON TRIM(LS.LOST_PROD_CODE) = TRIM(COD.TABLE_CODE)" +
                "   WHERE" +
                "     LS.EQUIP_NO = '" + equipmentNo + "'" +
                "     AND COD.TABLE_TYPE = 'LP'" +
                "     AND LS.LOST_PROD_CODE = '" + lostCode + "'" +
                "     AND (99999999 - LS.REV_STAT_DATE) = '" + startDate + "'" +
                "     AND LS.SHIFT = '" + shiftCode + "'" +
                "     AND LS.STOP_TIME = LPAD('" + startTime + "', 4, '0')" +
                "     AND LS.START_TIME = LPAD('" + endTime + "', 4, '0')";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }

        /// <summary>
        /// Consulta los eventos de la red industrial de PBV
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static string GetDownLostPbv(string startDate, string endDate)
        {
            var query = "" +                          
                           "WITH " +
                           "  SHIFT AS " +
                           "  ( " +
                           "    SELECT " +
                           "      CONVERT( DATETIME, '" + startDate + " 06:00:00', 20 ) STARTTIME, " +
                           "      CONVERT( DATETIME, '" + endDate + " 06:00:00', 20 ) ENDTIME " +
                           "  ) " +
                           "  , " +
                           "  PUSH AS " +
                           "  ( " +
                           "    SELECT " +
                           "      ROW_NUMBER( ) OVER( PARTITION BY PU.ASSETID ORDER BY PUSH.TIMESTAMP ASC ) ROWNUMBER, " +
                           "      PUSH.TIMESTAMP, " +
                           "      PU.ASSETID, " +
                           "      PUSH.PRODUCTIVEUNITSTATUSTAGVALUE PUS, " +
                           "      EVENTSEQUENCEID " +
                           "    FROM " +
                           "      PRODUCTIVEUNITSSTATUSHISTORY PUSH " +
                           "    INNER JOIN PRODUCTIVEUNITS PU " +
                           "    ON " +
                           "      PU.PRODUCTIVEUNITSID = PUSH.PRODUCTIVEUNITID " +
                           "    INNER JOIN PRODUCTIVEUNITSFUNCTIONSLIST PUFL " +
                           "    ON " +
                           "      PUFL.PRODUCTIVEUNITSID = PUSH.PRODUCTIVEUNITID " +
                           "    INNER JOIN DBO.IFIXSTATUSCODESHISTORY SCH " +
                           "    ON " +
                           "      PUFL.STATUSCODE_TAGNAME = SCH.FUNCTIONSTATUSTAGNAME " +
                           "    AND PUSH.TIMESTAMP        = SCH.TIMESTAMP " +
                           "  ) " +
                           "  , " +
                           "  PU_SEL AS " +
                           "  ( " +
                           "    SELECT " +
                           "      PUSH.TIMESTAMP   AS STARTTIME, " +
                           "      PUSH_1.TIMESTAMP AS ENDTIME, " +
                           "      PUSH.ASSETID     AS PUASSETID, " +
                           "      PUSH.EVENTSEQUENCEID " +
                           "    FROM " +
                           "      PUSH " +
                           "    INNER JOIN PUSH PUSH_1 " +
                           "    ON " +
                           "      PUSH.ASSETID         = PUSH_1.ASSETID " +
                           "    AND PUSH.ROWNUMBER + 1 = PUSH_1.ROWNUMBER " +
                           "    WHERE " +
                           "      PUSH.PUS                = 60000 " +
                           "    OR( PUSH.EVENTSEQUENCEID IS NOT NULL " +
                           "    AND PUSH.PUS              = 30000 ) " +
                           "  ) " +
                           "  , " +
                           "  SC AS " +
                           "  ( " +
                           "    SELECT " +
                           "      CASE SCC.ID " +
                           "        WHEN 13 " +
                           "        THEN 'LOST' " +
                           "        WHEN 15 " +
                           "        THEN 'LOST' " +
                           "        WHEN 14 " +
                           "        THEN 'DOWN' " +
                           "        WHEN 23 " +
                           "        THEN 'DOWN' " +
                           "        ELSE 'DOWN' " +
                           "      END TIPO, " +
                           "      SCC.ID, " +
                           "      SCC.DESCRIPTION " +
                           "    FROM " +
                           "      SCADARDB.DBO.STATUSCHANGECAUSE SCC " +
                           "    WHERE " +
                           "      SCC.PARENTID IS NULL " +
                           "    UNION ALL " +
                           "    SELECT " +
                           "      SC.TIPO, " +
                           "      SCC.ID, " +
                           "      SCC.DESCRIPTION " +
                           "    FROM " +
                           "      SC " +
                           "    INNER JOIN SCADARDB.DBO.STATUSCHANGECAUSE SCC " +
                           "    ON " +
                           "      SC.ID = SCC.PARENTID " +
                           "  ) " +
                           "  , " +
                           "  PU_EVENT AS " +
                           "  ( " +
                           "    SELECT " +
                           "      PUASSET.ASSETDESC EQUIPMENT, " +
                           "      REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) COMPONENT, " +
                           "      CASE " +
                           "        WHEN PU_SEL.STARTTIME < SHIFT.STARTTIME " +
                           "        THEN SHIFT.STARTTIME " +
                           "        ELSE PU_SEL.STARTTIME " +
                           "      END STARTTIME, " +
                           "      CASE " +
                           "        WHEN PU_SEL.ENDTIME > SHIFT.ENDTIME " +
                           "        THEN SHIFT.ENDTIME " +
                           "        ELSE PU_SEL.ENDTIME " +
                           "      END ENDTIME, " +
                           "      CASE SC.TIPO " +
                           "        WHEN 'DOWN' " +
                           "        THEN 'F' + REPLICATE( '0', 3 - LEN( ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SC.ID ), '    ' ) " +
                           "        WHEN 'LOST' " +
                           "        THEN 'L' + REPLICATE( '0', 3 - LEN( ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SC.ID ), '    ' ) " +
                           "        ELSE 'DW' " +
                           "      END FAILURE, " +
                           "      CASE " +
                           "        WHEN SYMPTOMS.SYMPTOMID IS NULL " +
                           "        THEN NULL " +
                           "        ELSE 'S' + REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, SYMPTOMS.SYMPTOMID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SYMPTOMS.SYMPTOMID ), '0' ) " +
                           "      END SYMPTOMID, " +
                           "      CASE " +
                           "        WHEN FAILUREASSET.ASSETTYPEID IS NULL " +
                           "        THEN NULL " +
                           "        ELSE 'P' + REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, FAILUREASSET.ASSETTYPEID ), '0' ) " +
                           "      END ASSETTYPEID, " +
                           "      CASE " +
                           "        WHEN SC.ID IS NULL " +
                           "        THEN NULL " +
                           "        ELSE 'C' + REPLICATE( '0', 4 - LEN( ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, SC.ID ), '0' ) " +
                           "      END STATUSCHANGEID, " +
                           "      CASE WHEN SC.TIPO = 'DOWN' " +
                           "        THEN 'EP' + REPLICATE( '0', 6 - LEN( ISNULL( CONVERT( VARCHAR, EH.EVENTSEQUENCEID ), '0' ) ) ) + ISNULL( CONVERT( VARCHAR, EH.EVENTSEQUENCEID ), '0' ) " +
                           "        ELSE NULL " +
                           "      END EVENT, " +
                           "      SC.TIPO, " +
                           "      SC.DESCRIPTION, " +
                           "      EH.COMMENT " +
                           "    FROM " +
                           "      PU_SEL " +
                           "    INNER JOIN SCADARDB.DBO.ASSETS PUASSET " +
                           "    ON " +
                           "      PU_SEL.PUASSETID = PUASSET.ASSETID " +
                           "    LEFT JOIN SCADARDB.DBO.EVENTSHISTORY EH " +
                           "    ON " +
                           "      PU_SEL.EVENTSEQUENCEID = EH.EVENTSEQUENCEID " +
                           "    LEFT JOIN SCADARDB.DBO.ASSETS FAILUREASSET " +
                           "    ON " +
                           "      EH.FAILEDASSETID = FAILUREASSET.ASSETID " +
                           "    LEFT JOIN SC " +
                           "    ON " +
                           "      EH.FAILEDASSETFAILUREMODEID = SC.ID " +
                           "    LEFT JOIN SYMPTOMS " +
                           "    ON " +
                           "      EH.SYMPTOMID = SYMPTOMS.SYMPTOMID " +
                           "    INNER JOIN SHIFT " +
                           "    ON " +
                           "      SHIFT.STARTTIME <= PU_SEL.ENDTIME " +
                           "    AND SHIFT.ENDTIME >= PU_SEL.STARTTIME " +
                           "  ) " +
                           "SELECT " +
                           "  PU_EVENT.EQUIPMENT, " +
                           "  PU_EVENT.COMPONENT, " +
                           "  '' COMP_MOD_CODE, " +
                           "  CONVERT( VARCHAR, PU_EVENT.STARTTIME, 112 ) STAR_DATE, " +
                           "  REPLACE( CONVERT( VARCHAR( 5 ), PU_EVENT.STARTTIME, 108 ), ':', '' ) STAR_TIME, " +
                           "  CONVERT( VARCHAR, PU_EVENT.ENDTIME, 112 ) FINISH_DATE, " +
                           "  REPLACE( CONVERT( VARCHAR( 5 ), PU_EVENT.ENDTIME, 108 ), ':', '' ) FINISH_TIME, " +
                           "  '' ELAPSED, " +
                           "  '' COLLECTION, " +
                           "  'A' SHIFT, " +
                           "  PU_EVENT.TIPO EVENT_TYPE, " +
                           "  PU_EVENT.FAILURE EVENT_CODE, " +
                           "  PU_EVENT.DESCRIPTION EVENT_DESC, " +
                           "  PU_EVENT.COMMENT, " +
                           "  PU_EVENT.EVENT, " +
                           "  PU_EVENT.SYMPTOMID, " +
                           "  PU_EVENT.ASSETTYPEID, " +
                           "  PU_EVENT.STATUSCHANGEID " +
                           "FROM " +
                           "  PU_EVENT ";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");
            
            return query;
        }
    }
}
