using System;
using System.Collections.Generic;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Utilities.Shifts;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse.Constants;
using SharedClassLibrary.Utilities;
using EllipseEquipmentClassLibrary;
using EllipseStdTextClassLibrary;
using EllipseWorkOrdersClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Screen = SharedClassLibrary.Ellipse.ScreenService;


namespace EllipseDownLostExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
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
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(SharedResources.Warning_DecimalSeparatorWarning, SharedResources.Warning_WarningUppercase);
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
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound} . {ex.Message}");
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
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLostPbv()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound} . {ex.Message}");
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
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    _ignoreDuplicate = false;
                    _thread = new Thread(CreateDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound} . {ex.Message}");
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
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _ignoreDuplicate = true;
                    _thread = new Thread(CreateDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound} . {ex.Message}");
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
                    var dr = MessageBox.Show(DlResources.Delete_Warning, DlResources.Delete_Title, MessageBoxButtons.YesNo);
                    if (dr != DialogResult.Yes)
                        return;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    _thread = new Thread(DeleteDownLost);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DeleteDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound} . {ex.Message}");
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
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01 || _excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _thread = new Thread(GenerateCollectionList);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(SharedResources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GenerateCollectionList()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound} . {ex.Message}");
            }
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
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName, 1);
                _cells.GetCell("B3").Value = "ICOR";

                var equipTypeList = new List<string> { SearchFieldCriteria.EquipmentReference.Value, SearchFieldCriteria.Egi.Value, SearchFieldCriteria.ListType.Value, SearchFieldCriteria.ProductiveUnit.Value };

                _cells.SetValidationList(_cells.GetCell("A4"), equipTypeList, ValidationSheetName, 2);
                _cells.GetCell("A4").Value = SearchFieldCriteria.EquipmentReference.Value;

                _cells.GetCell("A5").Value = SearchFieldCriteria.ListId.Value;
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
                var dataTypeList = new List<string> { "DOWN", "LOST", "DOWN & LOST" };

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
                _cells.GetCell(09, TitleRow01).AddComment(ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftCodes.DailyMorningCode + ": Dia 06-06");
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
                    ShiftCodes.HourToHourCode,
                    ShiftCodes.DailyZeroCode,
                    ShiftCodes.DailyMorningCode,
                    ShiftCodes.DayNightCode
                };
                _cells.SetValidationList(_cells.GetRange(09, TitleRow01 + 1, 09, TitleRow01 + 101), collectionList, ValidationSheetName, 4);

                var typeEvent = new List<string> { "DOWN", "LOST" };
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

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var dr = _eFunctions.GetQueryResult(Queries.GetDownTimeCodeListQuery(_eFunctions.DbReference, _eFunctions.DbLink));

                if (dr != null && !dr.IsClosed)
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

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                dr = _eFunctions.GetQueryResult(Queries.GetLostProdCodeListQuery(_eFunctions.DbReference, _eFunctions.DbLink));

                if (dr != null && !dr.IsClosed)
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
                _cells.GetCell(09, TitleRow01).AddComment(ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftCodes.DailyMorningCode + ": Dia 06-06");
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


                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01 - 1, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01 - 1, TitleRow01 + 1), TableName04);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();


                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_SheetHeaderError} . {ex.Message}");
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

                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName, 1);
                _cells.GetCell("B3").Value = "ICOR";

                var equipTypeList = new List<string> { SearchFieldCriteria.EquipmentReference.Value, SearchFieldCriteria.Egi.Value, SearchFieldCriteria.ListType.Value, SearchFieldCriteria.ProductiveUnit.Value };

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
                var dataTypeList = new List<string> { "DOWN", "LOST", "DOWN & LOST" };

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
                _cells.GetCell(09, TitleRow01).AddComment(ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftCodes.DailyMorningCode + ": Dia 06-06");
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
                    ShiftCodes.HourToHourCode,
                    ShiftCodes.DailyZeroCode,
                    ShiftCodes.DailyMorningCode,
                    ShiftCodes.DayNightCode
                };
                _cells.SetValidationList(_cells.GetRange(09, TitleRow01 + 1, 09, TitleRow01 + 101), collectionList, ValidationSheetName, 4);

                var typeEvent = new List<string> { "DOWN", "LOST" };
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

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var dr = _eFunctions.GetQueryResult(Queries.GetDownTimeCodeListQuery(_eFunctions.DbReference, _eFunctions.DbLink));

                if (dr != null && !dr.IsClosed)
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

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                dr = _eFunctions.GetQueryResult(Queries.GetLostProdCodeListQuery(_eFunctions.DbReference, _eFunctions.DbLink));

                if (dr != null && !dr.IsClosed)
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
                _cells.GetCell(09, TitleRow01).AddComment(ShiftCodes.HourToHourCode + ": Hora a Hora\n" + ShiftCodes.DayNightCode + ": Dia 06-18 y Noche 18-06 \n" + ShiftCodes.DailyZeroCode + ": Dia 00-24\n" + ShiftCodes.DailyMorningCode + ": Dia 06-06");
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
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_SheetHeaderError} . {ex.Message}");
            }
        }

        public void ReviewDownLostPbv()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var cells = new ExcelStyleCells(_excelApp, true);
                _cells.SetCursorWait();
                cells.ClearTableRange(TableName01);

                var startDate = _cells.GetNullIfTrimmedEmpty(cells.GetCell("D3").Value);
                var endDate = _cells.GetNullIfTrimmedEmpty(cells.GetCell("D4").Value);

                _eFunctions.SetDBSettings(Environments.ScadaRdb);
                var sqlQuery = Queries.GetDownLostPbv(startDate, endDate);

                var reader = _eFunctions.GetSqlQueryResult(sqlQuery);
                var currentRow = TitleRow01 + 1;

                if (reader == null || reader.IsClosed)
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
                    currentRow++;
                }
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLostPbv()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound}. {ex.Message}");
            }
            finally
            {
				_eFunctions.CloseConnection();
                _cells?.SetCursorDefault();
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
                var cells = new ExcelStyleCells(_excelApp, true);
                _cells.SetCursorWait();
                cells.ClearTableRange(TableName01);



                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var districtCode = cells.GetNullIfTrimmedEmpty(cells.GetCell("B3").Value);
                var equipType = cells.GetNullIfTrimmedEmpty(cells.GetCell("A4").Value);
                var equipRef = cells.GetNullIfTrimmedEmpty(cells.GetCell("B4").Value); //Equipment, EGI, PU, ListType
                var listId = cells.GetNullIfTrimmedEmpty(cells.GetCell("B5").Value);
                var startDate = cells.GetNullIfTrimmedEmpty(cells.GetCell("D3").Value);
                var endDate = cells.GetNullIfTrimmedEmpty(cells.GetCell("D4").Value);
                var dataType = cells.GetNullIfTrimmedEmpty(cells.GetCell("D5").Value);

                List<string> equipmentList;
                if (equipType.Equals(SearchFieldCriteria.EquipmentNo.Value) || equipType.Equals(SearchFieldCriteria.EquipmentReference.Value))
                    equipmentList = EquipmentActions.GetEquipmentList(_eFunctions, districtCode, equipRef);
                else if (equipType.Equals(SearchFieldCriteria.Egi.Value))
                    equipmentList = EquipmentActions.GetEgiEquipments(_eFunctions, equipRef);
                else if (equipType.Equals(SearchFieldCriteria.ListType.Value))
                    equipmentList = EquipmentActions.GetListEquipments(_eFunctions, equipRef, listId);
                else if (equipType.Equals(SearchFieldCriteria.ProductiveUnit.Value))
                    equipmentList = EquipmentActions.GetProductiveUnitEquipments(_eFunctions, districtCode, equipRef);
                else
                    equipmentList = EquipmentActions.GetEquipmentList(_eFunctions, districtCode, equipRef);

                var i = TitleRow01 + 1;

                foreach (var equip in equipmentList)
                {
                    if (dataType.Equals("DOWN") || dataType.Equals("DOWN & LOST"))
                    {
                        var sqlDownQuery = Queries.GetEquipmentDownQuery(_eFunctions.DbReference, _eFunctions.DbLink, equip,
                            startDate, endDate);
                        var ddr = _eFunctions.GetQueryResult(sqlDownQuery);

                        if (ddr != null && !ddr.IsClosed)
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
                                if(!string.IsNullOrWhiteSpace(ddr["WORK_ORDER"].ToString()))
                                    cells.GetCell(ResultColumn01, i).Value = "Estado OT: " + cells.GetEmptyIfNull(WoStatusList.GetStatusName(ddr["WO_STATUS_M"].ToString()));
                                
                                cells.GetCell(01, i).Select();
                                i++;
                            }
                        }
                    }
                    if (dataType.Equals("LOST") || dataType.Equals("DOWN & LOST"))
                    {
                        var lostSqlQuery = Queries.GetEquipmentLostQuery(_eFunctions.DbReference,
                            _eFunctions.DbLink, equip, startDate, endDate);
                        var ddr = _eFunctions.GetQueryResult(lostSqlQuery);

                        if (ddr != null && !ddr.IsClosed)
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
                }
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ReviewDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_ErrorFound}. {ex.Message}");
            }
            finally
            {
				_eFunctions.CloseConnection();
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
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var cells = new ExcelStyleCells(_excelApp, true);

                var cellCollection = new ExcelStyleCells(_excelApp, SheetName04);

                var i = TitleRow01 + 1;

                var resultColumn01 = (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01) ? ResultColumnP01 : ResultColumn01;

                while ("" + cells.GetCell(1, i).Value != "")
                {
                    try
                    {
                        //ScreenService Opción en reemplazo de los servicios
                        var opSheet = new Screen.OperationContext
                        {
                            district = _frmAuth.EllipseDstrct,
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

                        if ((eventType.ToUpper().Equals("DOWN") || eventType.ToUpper().Equals("D")) && !string.IsNullOrWhiteSpace(woComment))
                            woComment = MyUtilities.GetCodeKey(woComment);

                        woComment = string.IsNullOrWhiteSpace(woComment) ? null : woComment;
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
                        if (collection == ShiftCodes.HourToHourCode ||
                            //hora a hora (Ej. 00-01, 01-02, ..., 22-23, 23-24
                            collection == ShiftCodes.DailyZeroCode || //diaria de 00-24
                            collection == ShiftCodes.DailyMorningCode || //diaria de 06-06
                            collection == ShiftCodes.DayNightCode) //dia 06-18 y noche de 18-06
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
                                MyUtilities.DateTime.GetSlots(GetTurnShifts(collection), startEvent, endEvent).ToArray();
                            

                            ldObject = new LostDownObject[shiftArray.Length];

                            for (var j = 0; j < shiftArray.Length; j++)
                            {
                                var dateString = MyUtilities.ToString(shiftArray[j].GetDate());
                                var startTimeString = MyUtilities.ToString(shiftArray[j].GetStartDateTime().TimeOfDay, MyUtilities.DateTime.TimeHHMM);
                                var endTimeString = MyUtilities.ToString(shiftArray[j].GetEndDateTime().TimeOfDay, MyUtilities.DateTime.TimeHHMM);

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
                        Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                    }
                    finally
                    {
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
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
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var cells = new ExcelStyleCells(_excelApp, true);
                var cellCollection = new ExcelStyleCells(_excelApp, SheetName04);
                var resultColumn01 = (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01) ? ResultColumnP01 : ResultColumn01;
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


                        if (collection == ShiftCodes.HourToHourCode ||
                            //hora a hora (Ej. 00-01, 01-02, ..., 22-23, 23-24
                            collection == ShiftCodes.DailyZeroCode || //diaria de 00-24
                            collection == ShiftCodes.DailyMorningCode || //diaria de 06-06
                            collection == ShiftCodes.DayNightCode) //dia 06-18 y noche de 18-06
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
                            var shiftArray = MyUtilities.DateTime.GetSlots(GetTurnShifts(collection), startEvent, endEvent).ToArray();

                            var ldObject = new LostDownObject[shiftArray.Length];

                            for (var j = 0; j < shiftArray.Length; j++)
                            {

                                var dateString = MyUtilities.ToString(shiftArray[j].GetDate());
                                var startTimeString = MyUtilities.ToString(shiftArray[j].GetStartDateTime().TimeOfDay, "hhmm");
                                var endTimeString = MyUtilities.ToString(shiftArray[j].GetEndDateTime().TimeOfDay, MyUtilities.DateTime.TimeHHMM);

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

                        cells.GetCell(resultColumn01, i).Value = DlResources.Field_CollectionUppercase;
                        cells.GetCell(resultColumn01, i).Style = StyleConstants.Success;
                        cells.GetCell(resultColumn01, i).Select();
                    }
                    catch (Exception ex)
                    {
                        cells.GetCell(resultColumn01, i).Value = SharedResources.Error_ErrorUppercase + ":" + ex.Message;
                        cells.GetCell(resultColumn01, i).Style = StyleConstants.Error;
                        cells.GetCell(resultColumn01, i).Select();
                        Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                    }
                    finally
                    {
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }
        private void AddDownLostToTableCollection(LostDownObject lostDownObject, string eventType, ExcelStyleCells cells)
        {
            var tableName = TableName04;
            var titleRow = TitleRow01;
            //Escribo el objeto de colecction en la tabla collection
            var tableRange = cells.GetRange(tableName);
            var row = tableRange.ListObject.ListColumns[1].Range.Row + tableRange.ListObject.ListRows.Count + 1;
            
            //para controlar la escritura en table si está vacía
            if (string.IsNullOrWhiteSpace(cells.GetCell(01, titleRow + 1).Value2))
                row = titleRow + 1;
            //
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
        /// <param name="screenService">Screen.ScreenService: Servicio de Screen Service a utilizar</param>
        /// <param name="ldObject">LostDownObject[] : Arreglo de objetos a adicionar para Down</param>
        /// <param name="ignoreDuplicate">bool: true para ignorar duplicado para el cargue de colección</param>
        public void CreateDownRegister(Screen.OperationContext opContext, Screen.ScreenService screenService, LostDownObject[] ldObject, bool ignoreDuplicate = false)
        {
            foreach (var down in ldObject)
            {
                try
                {
                    screenService.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
                    _eFunctions.RevertOperation(opContext, screenService);
                    //ejecutamos el programa
                    var reply = screenService.executeScreen(opContext, "MSO420");
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
                    reply = screenService.submit(opContext, request);

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
                            request = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                            screenService.submit(opContext, request);
                            replyFields = new ArrayScreenNameValue(reply.screenFields);
                        }

                        if (down.WoEvent != null && WorkOrderActions.FetchWorkOrder(_eFunctions, "", down.WoEvent) == null)
                        {
                            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
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
                                _woDownOriginator = InputBox.GetValue(DlResources.Field_WorkOrderCamelcase, DlResources.Input_InputOriginatorId, DlResources.Field_UsernameUppercase);
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
                        var woOrder = down.WoEvent ?? down.WoComment;
                        if(!string.IsNullOrWhiteSpace(woOrder))
                            arrayFields.Add("WORK_ORDER1I" + k, down.WoEvent ?? down.WoComment);//si el WoEvent es nulo es porque no es un evento de SCADA PBV si no un Down regular
                        arrayFields.Add("COMP_CODE1I" + k, down.CompCode);
                        arrayFields.Add("MODIFIER1I" + k, down.CompModCode);

                        request = new Screen.ScreenSubmitRequestDTO
                        {
                            screenFields = arrayFields.ToArray(),
                            screenKey = "1"
                        };
                        reply = screenService.submit(opContext, request);

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
                            reply = screenService.submit(opContext, request);


                        }
                        if (reply != null && (_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM420A"))
                            throw new ArgumentException(reply.message);


                    }
                    else if (reply != null) throw new Exception(reply.message);
                    else throw new Exception(DlResources.Error_NoServerResponse);
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
                    proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
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
                            request = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
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
                        var dr = _eFunctions.GetQueryResult(Queries.GetSingleLostQuery(_eFunctions.DbReference, _eFunctions.DbLink, lost.EquipNo, lost.EventCode, lost.Date, lost.ShiftCode, lost.StartTime, lost.FinishTime));
                        var stdTextId = "";
                        if (dr != null && !dr.IsClosed)
                            while (dr.Read())
                                stdTextId = _cells.GetEmptyIfNull(dr["STD_KEY"].ToString());
                        else
                            throw new Exception(DlResources.Error_NoLostForComment);
                        dr.Close();

                        var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
                        stdTextId = "LP" + stdTextId;
                        var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

                        var textResult = StdText.SetText(urlService, StdText.GetCustomOpContext(district, _frmAuth.EllipsePost, 100, false), stdTextId, lost.WoComment);

                        if (!textResult)
                            throw new KeyNotFoundException(DlResources.Error_LostCommentCreationFailed);
                    }
                    else if (reply != null) throw new Exception(reply.message);
                    else throw new Exception(DlResources.Error_NoServerResponse);
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
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                var cells = new ExcelStyleCells(_excelApp, true);

                var i = TitleRow01 + 1;
                
                var resultColumn01 = (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameP01) ? ResultColumnP01 : ResultColumn01;

                while ("" + cells.GetCell(1, i).Value != "")
                {
                    try
                    {

                        //ScreenService Opción en reemplazo de los servicios
                        var opSheet = new Screen.OperationContext
                        {
                            district = _frmAuth.EllipseDstrct,
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
                            cells.GetCell(resultColumn01, i).Value = SharedResources.Results_Deleted;
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Success;
                            cells.GetCell(resultColumn01, i).Select();
                        }
                        else
                        {
                            cells.GetCell(resultColumn01, i).Value = DlResources.Result_NoActionExecuted;
                            cells.GetCell(resultColumn01, i).Style = StyleConstants.Warning;
                            cells.GetCell(resultColumn01, i).Select();
                        }
                    }
                    catch (Exception ex)
                    {
                        cells.GetCell(resultColumn01, i).Value = SharedResources.Error_ErrorUppercase + ": " + ex.Message;
                        cells.GetCell(resultColumn01, i).Style = StyleConstants.Error;
                        Debugger.LogError("RibbonEllipse:DeleteDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
                    }
                    finally
                    {
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:CreateDownLost()", "\n\r" + SharedResources.Debugging_Message + ":" + ex.Message + "\n\r" + SharedResources.Debugging_Source + ":" + ex.Source + "\n\r" + SharedResources.Debugging_StackTrace + ":" + ex.StackTrace);
            }
            finally
            {
                _cells?.SetCursorDefault();
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
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
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

                if (k <= 10) continue;
                k = 1;
                //envíe a la siguiente pantalla
                request = new Screen.ScreenSubmitRequestDTO {screenKey = "1"};
                reply = proxySheet.submit(opContext, request);
				replyFields = new ArrayScreenNameValue(reply.screenFields);
            }

            if (reply.mapName != "MSM420A")
                throw new ArgumentException(SharedResources.Error_ItemNotFound);

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
            throw new Exception(DlResources.Error_NoServerResponse);
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
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
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
                request = new Screen.ScreenSubmitRequestDTO {screenKey = "1"};
                reply = proxySheet.submit(opContext, request);
				replyFields = new ArrayScreenNameValue(reply.screenFields);
            }

            if (reply.mapName != "MSM470A")
                throw new ArgumentException(SharedResources.Error_ItemNotFound);

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
            throw new Exception(DlResources.Error_NoServerResponse);
        }

        /// <summary>
        /// Obtiene un arreglo de tipo ShiftSlot [] con la información de los turnos correspondientes a una colección de un tipo de turno (Ej: Colleción D/N tiene los turnos D de 0600 a 1800 y N de 1800 a 0600; Collección HH tiene los turnos 01 de 0000 a 0100, ..., y 24 de 2300 a 0000)
        /// </summary>
        /// <param name="shiftPeriodCode">Código de colección del Turno</param>
        /// <returns></returns>
        private static Slot[] GetTurnShifts(string shiftPeriodCode)
        {
            if (shiftPeriodCode.Equals(ShiftCodes.HourToHourCode))
                return ShiftPeriods.GetHourToHourSlots();
            if (shiftPeriodCode.Equals(ShiftCodes.DailyMorningCode))
                return ShiftPeriods.GetDailyMorningSlots();
            if (shiftPeriodCode.Equals(ShiftCodes.DailyZeroCode))
                return ShiftPeriods.GetDailyZeroSlots();
            // ReSharper disable once ConvertIfStatementToReturnStatement
            if (shiftPeriodCode.Equals(ShiftCodes.DayNightCode))
                return ShiftPeriods.GetDailyNightSlots();
            return null;
        }

        
        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread == null || !_thread.IsAlive) return;
                _thread.Abort();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show($@"{SharedResources.Error_ThreadStopped} . {ex.Message}");
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }

    
}
