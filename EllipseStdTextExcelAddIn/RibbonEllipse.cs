using System;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Vsto.Excel;
using SharedClassLibrary.Utilities;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel; 
using EllipseStdTextClassLibrary;
using EllipseReferenceCodesClassLibrary;
using EllipseDocumentReferenceClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;

namespace EllipseStdTextExcelAddIn
{
    public partial class RibbonEllipse
    {

        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Excel.Application _excelApp;
        private const string SheetName01 = "StdText";
        private const string SheetName02 = "ReferenceCodes";
        private const string SheetName03 = "DocReferences";
        private const int TitleRow01 = 5;
        private const int TitleRow02 = 5;
        private const int TitleRow03 = 5;
        private const int ResultColumn01 = 6;
        private const int ResultColumn02 = 10;
        private const int ResultColumn03 = 14;
        private const string TableName01 = "StdTextTable";
        private const string TableName02 = "RefCodeTable";
        private const string TableName03 = "DocRefTable";

        private const string ValidationSheetName = "ValidationData";

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

        private void btnGetHeaderAndText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => GetStdText(true, true));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:GetStdText(true, true)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateHeaderAndText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => SetStdText(true, true));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:SetStdText(true, true)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            
        }

        private void btnGetHeaderOnly_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => GetStdText(true, false));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:GetStdText(true, false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnSetHeaderOnly_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => SetStdText(true, false));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:SetStdText(true, false)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnGetTextOnly_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => GetStdText(false, true));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:GetStdText(false, true)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnSetTextOnly_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(() => SetStdText(false, true));
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:SetStdText(false, true)", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
               

                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                //CONSTRUYO LA HOJA 1
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STD TEXT - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);


                _cells.GetCell(1, TitleRow01).Value = "TYPE[2]";
                _cells.GetCell(1, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(2, TitleRow01).Value = "DISTRICT[4]";
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, TitleRow01).Value = "ID[8]";
                _cells.GetCell(3, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(4, TitleRow01).Value = "HEADER";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, TitleRow01).Value = "TEXT";
                _cells.GetCell(5, TitleRow01).Style = StyleConstants.TitleRequired;

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow01+1 , ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 2 - REFERENCE CODES
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "REFERENCE CODES - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);


                _cells.GetCell(1, TitleRow02).Value = "ENT. TYPE [3]";
                _cells.GetCell(1, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow02).Value = "ENT. VALUE 1";
                _cells.GetCell(2, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(3, TitleRow02).Value = "ENT. VALUE 2";
                _cells.GetCell(3, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, TitleRow02).Value = "ENT. NO";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow02).Value = "ENT. SEQ";
                _cells.GetCell(5, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, TitleRow02).Value = "REF CODE";
                _cells.GetCell(6, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(7, TitleRow02).Value = "ST FLAG";
                _cells.GetCell(7, TitleRow02).Style = StyleConstants.TitleInformation;
                _cells.GetCell(8, TitleRow02).Value = "STDTEXT ID";
                _cells.GetCell(8, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(9, TitleRow02).Value = "STDTEXT TEXT";
                _cells.GetCell(9, TitleRow02).Style = StyleConstants.TitleOptional;

                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
               

                //CONSTRUYO LA HOJA 3 - DOCUMENT REFERENCES
                _excelApp.ActiveWorkbook.Sheets[3].Select(Type.Missing);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "DOCUMENT REFERENCES - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);


                _cells.GetCell(1, TitleRow03).Value = "Reference Type";
                _cells.GetCell(1, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, TitleRow03).Value = "Reference No";
                _cells.GetCell(2, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(3, TitleRow03).Value = "Reference Other";
                _cells.GetCell(3, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, TitleRow03).Value = "Document No";
                _cells.GetCell(4, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, TitleRow03).Value = "Document Reference";
                _cells.GetCell(5, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, TitleRow03).Value = "Prefix";
                _cells.GetCell(6, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, TitleRow03).Value = "Document Type";
                _cells.GetCell(7, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(8, TitleRow03).Value = "Version Status";
                _cells.GetCell(8, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(9, TitleRow03).Value = "Version Type";
                _cells.GetCell(9, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow03).Value = "Version No";
                _cells.GetCell(10, TitleRow03).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, TitleRow03).Value = "Description";
                _cells.GetCell(11, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(12, TitleRow03).Value = "Electronic Reference";
                _cells.GetCell(12, TitleRow03).Style = StyleConstants.TitleRequired;
                _cells.GetCell(13, TitleRow03).Value = "Electronic Type";
                _cells.GetCell(13, TitleRow03).Style = StyleConstants.TitleOptional;



                _cells.GetCell(ResultColumn03, TitleRow03).Value = "RESULTADO";
                _cells.GetCell(ResultColumn03, TitleRow03).Style = StyleConstants.TitleResult;

                var refTypeCodes = _eFunctions.GetItemCodesString("DOLT");
                var docTypeCodes = _eFunctions.GetItemCodesString("DO");
                var versionTypeCodes = _eFunctions.GetItemCodesString("VT");
                var versionStatusCodes = _eFunctions.GetItemCodesString("DOVS");
                var elecTypeCodes = _eFunctions.GetItemCodesString("DOET");

                _cells.SetValidationList(_cells.GetCell(1, TitleRow03 + 1), refTypeCodes, ValidationSheetName, 1, false);
                _cells.SetValidationList(_cells.GetCell(7, TitleRow03 + 1), docTypeCodes, ValidationSheetName, 2, false);
                _cells.SetValidationList(_cells.GetCell(8, TitleRow03 + 1), versionStatusCodes, ValidationSheetName, 3, false);
                _cells.SetValidationList(_cells.GetCell(9, TitleRow03 + 1), versionTypeCodes, ValidationSheetName, 4, false);
                _cells.SetValidationList(_cells.GetCell(13, TitleRow03 + 1), elecTypeCodes, ValidationSheetName, 5, false);

                _cells.GetRange(1, TitleRow03 + 1, ResultColumn03, TitleRow03 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumn03, TitleRow03 + 1), TableName03);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show($@"{SharedResources.Error_SheetHeaderError} {ex.Message}");
            }
        }

        private void GetStdText(bool getHeader, bool getText)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var position = _frmAuth.EllipsePost;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value))
            {
                try
                {
                    var stdTextId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var stdTextOpc = StdText.GetStdTextOpContext(districtCode, position, 100, Debugger.DebugWarnings);
                    if (getHeader)
                    {
                        
                        string headerText = StdText.GetHeader(urlService, stdTextOpc, stdTextId);
                        _cells.GetCell(4, i).Value = "" + headerText;
                    }

                    if (getText)
                    {
                        string bodyText = StdText.GetText(urlService, stdTextOpc, stdTextId);
                        _cells.GetCell(5, i).Value = "" + bodyText;
                    }

                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:GetStdText(bool, bool)", ex.Message);
                }
                finally
                {
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }

        private void SetStdText(bool setHeader, bool setText)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            var position = _frmAuth.EllipsePost;
            var stdTextOpc = StdText.GetStdTextOpContext(districtCode, position, 100, Debugger.DebugWarnings);
            var stdTextCustomOpc = StdText.GetCustomOpContext(districtCode, position, 100, Debugger.DebugWarnings);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value))
            {
                try
                {
                    var stdTextId = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value) +
                                    _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);

                    var resultHeader = true;
                    var resultBody = true;

                    if (setHeader)
                    {
                        var headerText = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);
                        resultHeader = StdText.SetHeader(urlService, stdTextOpc, stdTextId, headerText);
                    }

                    if (setText)
                    {
                        var bodyText = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                        resultBody = StdText.SetText(urlService, stdTextCustomOpc, stdTextId, bodyText);
                    }

                    if (resultHeader && resultBody)
                    {
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                        _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                    }
                    else
                    {
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, i).Value = "ERROR";
                    }

                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:GetStdText(bool, bool)", ex.Message);
                }
                finally
                {
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }

        private void ReviewRefCodesList()
        {

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow02 + 1;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            _eFunctions.SetConnectionPoolingType(false);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {

                    var entityType = "" + _cells.GetCell(1, i).Value;
                    var entityValue = "" + _cells.GetCell(2, i).Value + _cells.GetCell(3, i).Value;
                    var refNum = "" + _cells.GetCell(4, i).Value;
                    var seqNum = "" + _cells.GetCell(5, i).Value;

                    refNum = string.IsNullOrWhiteSpace(refNum) ? "001" : refNum.PadLeft(3, '0');
                    seqNum = string.IsNullOrWhiteSpace(seqNum) ? "001" : seqNum.PadLeft(3, '0'); 

                    if (entityType.Equals("WRQ"))
                        entityValue = entityValue.PadLeft(12, '0');

                    var refItem = ReferenceCodeActions.FetchReferenceCodeItem(_eFunctions, urlService, rcOpContext, entityType, entityValue, refNum, seqNum);
                    //GENERAL
                    _cells.GetCell(6, i).Value = "'" + refItem.RefCode;
                    _cells.GetCell(7, i).Value = "'" + refItem.StdTextFlag;
                    _cells.GetCell(8, i).Value = "'" + refItem.StdtxtId;
                    _cells.GetCell(9, i).Value = "'" + refItem.StdText;

                    _cells.GetCell(ResultColumn02, i).Value = "CONSULTADO";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _eFunctions.SetConnectionPoolingType(true);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }

        private void UpdateRefCodesList()
        {

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var rcOpContext = ReferenceCodeActions.GetRefCodesOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow02 + 1;

            //Se encuentran problemas de implementación, debido a un comportamiento irregular del ODP en Windows. 
            //Las conexiones cerradas (EllipseFunctions.Close()) vuelven a la piscina (pool) de conexiones por un tiempo antes 
            //de ser completamente Cerradas (Close) y Dispuestas (Dispose), lo que ocasiona un desbordamiento del
            //número máximo de conexiones en el pool (100) y la nueva conexión alcanza el tiempo de espera (timeout) antes de
            //entrar en la cola del pool de conexiones arrojando un error 'Pooled Connection Request Timed Out'.
            //Para solucionarlo se fuerza el string de conexiones para que no genere una conexión que entre al pool.
            //Esto implica mayor tiempo de ejecución pero evita la excepción por el desbordamiento y tiempo de espera
            _eFunctions.SetConnectionPoolingType(false);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {

                    var entityType = "" + _cells.GetCell(1, i).Value;
                    var entityValue = "" + _cells.GetCell(2, i).Value + _cells.GetCell(3, i).Value;
                    var refNum = "" + _cells.GetCell(4, i).Value;
                    var seqNum = "" + _cells.GetCell(5, i).Value;
                    var refCode = "" + _cells.GetCell(6, i).Value;
                    var stdTextId = "" + _cells.GetCell(8, i).Value;
                    var stdText = "" + _cells.GetCell(9, i).Value;
                    refNum = string.IsNullOrWhiteSpace(refNum) ? "001" : refNum.PadLeft(3, '0');
                    seqNum = string.IsNullOrWhiteSpace(seqNum) ? "001" : seqNum.PadLeft(3, '0'); 

                    if (entityType.Equals("WRQ"))
                        entityValue = entityValue.PadLeft(12, '0');

                    var refItem = new ReferenceCodeItem(entityType, entityValue, refNum, seqNum, refCode, stdTextId, stdText);
                    ReferenceCodeActions.ModifyRefCode(_eFunctions, urlService, rcOpContext, refItem);

                    _cells.GetCell(ResultColumn02, i).Value = "ACTUALIZADO";
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _eFunctions.SetConnectionPoolingType(true);
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells?.SetCursorDefault();
        }

        private void btnCleanTable_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                _cells.ClearTableRange(TableName01);
            else if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                _cells.ClearTableRange(TableName02);
        }

        private void btnReviewRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
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

        private void btnUpdateRefCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName02))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateRefCodesList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateRefCodesList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnCreateDocRef_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(CreateDocumentReference);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateDocumentReference()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnLinkDocument_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(LinkDocumentReference);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateDocumentReference()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnUpdateDocument_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateDocumentReference);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateDocumentReference()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnDeleteReference_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName03))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(DeleteDocumentReference);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateDocumentReference()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void CreateDocumentReference()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var drOpContext = DocumentReferenceActions.GetDocRefOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow03 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var refType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    var refNo = _cells.GetCell(2, i).Value;
                    var refOther = _cells.GetCell(3, i).Value;
                    var docNo = _cells.GetCell(4, i).Value;
                    var docRef = _cells.GetCell(5, i).Value;
                    var refPrefix = _cells.GetCell(6, i).Value;
                    var docType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    var versionType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    var versionStatus = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value));
                    var versionNo = _cells.GetCell(10, i).Value;
                    var docName = _cells.GetCell(11, i).Value;
                    var elecRef = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value));
                    var elecType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value));

                    var item = new DocumentReferenceItem();
                    item.District = string.IsNullOrWhiteSpace(district) ? null : district;
                    item.DocRefType = string.IsNullOrWhiteSpace(refType) ? null : refType;
                    item.DocReference = string.IsNullOrWhiteSpace(refNo) ? null : refNo;
                    item.DocRefOther = string.IsNullOrWhiteSpace(refOther) ? null : refOther;
                    item.DocumentNo = string.IsNullOrWhiteSpace(docNo) ? null : docNo;
                    item.DocumentRef = string.IsNullOrWhiteSpace(docRef) ? null : docRef;
                    item.DocPrefix = string.IsNullOrWhiteSpace(refPrefix) ? null : refPrefix;
                    item.DocumentType = string.IsNullOrWhiteSpace(docType) ? null : docType;
                    item.VerType = string.IsNullOrWhiteSpace(versionType) ? null : versionType;
                    item.VerStatus = string.IsNullOrWhiteSpace(versionStatus) ? null : versionStatus;
                    item.DocVerNo = string.IsNullOrWhiteSpace(versionNo) ? null : versionNo;
                    item.DocumentName1 = string.IsNullOrWhiteSpace(docName) ? null : docName;
                    item.ElecRef = string.IsNullOrWhiteSpace(elecRef) ? null : elecRef;
                    item.ElecType = string.IsNullOrWhiteSpace(elecType) ? null : elecType;

                    var reply = DocumentReferenceActions.CreateDocument(urlService, drOpContext, item);
                    var successMessage = "CREADO";

                    if (reply.errors != null && reply.errors.Length > 0)
                    {
                        string error = "";
                        foreach (var err in reply.errors)
                            error = error + err.messageText + "\n";

                        if (error.Contains("DOCUMENT ALREADY EXISTS"))
                        {
                            reply = DocumentReferenceActions.LinkDocument(urlService, drOpContext, item);
                            successMessage = "VINCULADO. DOCUMENTO YA EXISTENTE";

                            //revalido el error
                            if (reply.errors != null && reply.errors.Length > 0)
                            {
                                error = "";
                                foreach (var err in reply.errors)
                                    error = error + err.messageText + "\n";
                                throw new Exception(error);
                            }
                        }
                        else
                            throw new Exception(error);
                    }

                    var newItem = new DocumentReferenceItem(reply.documentReferenceDTO);

                    _cells.GetCell(1, i).Value = "" + newItem.DocRefType;
                    _cells.GetCell(2, i).Value = "" + newItem.DocReference;
                    _cells.GetCell(3, i).Value = "" + newItem.DocRefOther;
                    _cells.GetCell(4, i).Value = "" + newItem.DocumentNo;
                    _cells.GetCell(5, i).Value = "" + newItem.DocumentRef;
                    _cells.GetCell(6, i).Value = "" + newItem.DocPrefix;
                    _cells.GetCell(7, i).Value = "" + newItem.DocumentType;
                    _cells.GetCell(8, i).Value = "" + newItem.VerStatus;
                    _cells.GetCell(9, i).Value = "" + newItem.VerType;
                    _cells.GetCell(10, i).Value = "" + newItem.DocVerNo;
                    _cells.GetCell(11, i).Value = "" + newItem.DocumentName1;
                    _cells.GetCell(12, i).Value = "" + newItem.ElecRef;
                    _cells.GetCell(13, i).Value = "" + newItem.ElecType;

                    _cells.GetCell(ResultColumn03, i).Value = successMessage;
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }

            _cells?.SetCursorDefault();
        }

        private void LinkDocumentReference()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var drOpContext = DocumentReferenceActions.GetDocRefOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow03 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var refType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    var refNo = _cells.GetCell(2, i).Value;
                    var refOther = _cells.GetCell(3, i).Value;
                    var docNo = _cells.GetCell(4, i).Value;
                    var docRef = _cells.GetCell(5, i).Value;
                    var refPrefix = _cells.GetCell(6, i).Value;
                    var docType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    var versionType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    var versionStatus = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value));
                    var versionNo = _cells.GetCell(10, i).Value;
                    var docName = _cells.GetCell(11, i).Value;
                    var elecRef = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value));
                    var elecType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value));

                    var item = new DocumentReferenceItem();
                    item.District = string.IsNullOrWhiteSpace(district) ? null : district;
                    item.DocRefType = string.IsNullOrWhiteSpace(refType) ? null : refType;
                    item.DocReference = string.IsNullOrWhiteSpace(refNo) ? null : refNo;
                    item.DocRefOther = string.IsNullOrWhiteSpace(refOther) ? null : refOther;
                    item.DocumentNo = string.IsNullOrWhiteSpace(docNo) ? null : docNo;
                    item.DocumentRef = string.IsNullOrWhiteSpace(docRef) ? null : docRef;
                    item.DocPrefix = string.IsNullOrWhiteSpace(refPrefix) ? null : refPrefix;
                    item.DocumentType = string.IsNullOrWhiteSpace(docType) ? null : docType;
                    item.VerType = string.IsNullOrWhiteSpace(versionType) ? null : versionType;
                    item.VerStatus = string.IsNullOrWhiteSpace(versionStatus) ? null : versionStatus;
                    item.DocVerNo = string.IsNullOrWhiteSpace(versionNo) ? null : versionNo;
                    item.DocumentName1 = string.IsNullOrWhiteSpace(docName) ? null : docName;
                    item.ElecRef = string.IsNullOrWhiteSpace(elecRef) ? null : elecRef;
                    item.ElecType = string.IsNullOrWhiteSpace(elecType) ? null : elecType;

                    var reply = DocumentReferenceActions.LinkDocument(urlService, drOpContext, item);

                    if(reply.errors != null && reply.errors.Length > 0)
                    {
                        string error = "";
                        foreach(var err in reply.errors)
                        {
                            error = error + err.messageText + "\n";
                        }
                        throw new Exception(error);
                    }
                    var newItem = new DocumentReferenceItem(reply.documentReferenceDTO);

                    _cells.GetCell(1, i).Value = "" + newItem.DocRefType;
                    _cells.GetCell(2, i).Value = "" + newItem.DocReference;
                    _cells.GetCell(3, i).Value = "" + newItem.DocRefOther;
                    _cells.GetCell(4, i).Value = "" + newItem.DocumentNo;
                    _cells.GetCell(5, i).Value = "" + newItem.DocumentRef;
                    _cells.GetCell(6, i).Value = "" + newItem.DocPrefix;
                    _cells.GetCell(7, i).Value = "" + newItem.DocumentType;
                    _cells.GetCell(8, i).Value = "" + newItem.VerStatus;
                    _cells.GetCell(9, i).Value = "" + newItem.VerType;
                    _cells.GetCell(10, i).Value = "" + newItem.DocVerNo;
                    _cells.GetCell(11, i).Value = "" + newItem.DocumentName1;
                    _cells.GetCell(12, i).Value = "" + newItem.ElecRef;
                    _cells.GetCell(13, i).Value = "" + newItem.ElecType;


                    _cells.GetCell(ResultColumn03, i).Value = "VINCULADO";
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:LinkDocumentReference()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }

            _cells?.SetCursorDefault();
        }

        private void UpdateDocumentReference()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var drOpContext = DocumentReferenceActions.GetDocRefOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow03 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var refType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    var refNo = _cells.GetCell(2, i).Value;
                    var refOther = _cells.GetCell(3, i).Value;
                    var docNo = _cells.GetCell(4, i).Value;
                    var docRef = _cells.GetCell(5, i).Value;
                    var refPrefix = _cells.GetCell(6, i).Value;
                    var docType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    var versionType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    var versionStatus = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value));
                    var versionNo = _cells.GetCell(10, i).Value;
                    var docName = _cells.GetCell(11, i).Value;
                    var elecRef = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value));
                    var elecType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value));

                    var item = new DocumentReferenceItem();
                    item.District = string.IsNullOrWhiteSpace(district) ? null : district;
                    item.DocRefType = string.IsNullOrWhiteSpace(refType) ? null : refType;
                    item.DocReference = string.IsNullOrWhiteSpace(refNo) ? null : refNo;
                    item.DocRefOther = string.IsNullOrWhiteSpace(refOther) ? null : refOther;
                    item.DocumentNo = string.IsNullOrWhiteSpace(docNo) ? null : docNo;
                    item.DocumentRef = string.IsNullOrWhiteSpace(docRef) ? null : docRef;
                    item.DocPrefix = string.IsNullOrWhiteSpace(refPrefix) ? null : refPrefix;
                    item.DocumentType = string.IsNullOrWhiteSpace(docType) ? null : docType;
                    item.VerType = string.IsNullOrWhiteSpace(versionType) ? null : versionType;
                    item.VerStatus = string.IsNullOrWhiteSpace(versionStatus) ? null : versionStatus;
                    item.DocVerNo = string.IsNullOrWhiteSpace(versionNo) ? null : versionNo;
                    item.DocumentName1 = string.IsNullOrWhiteSpace(docName) ? null : docName;
                    item.ElecRef = string.IsNullOrWhiteSpace(elecRef) ? null : elecRef;
                    item.ElecType = string.IsNullOrWhiteSpace(elecType) ? null : elecType;

                    var reply = DocumentReferenceActions.UpdateDocument(urlService, drOpContext, item);

                    if (reply.errors != null && reply.errors.Length > 0)
                    {
                        string error = "";
                        foreach (var err in reply.errors)
                        {
                            error = error + err.messageText + "\n";
                        }
                        throw new Exception(error);
                    }
                    var newItem = new DocumentReferenceItem(reply.documentReferenceDTO);

                    _cells.GetCell(1, i).Value = "" + newItem.DocRefType;
                    _cells.GetCell(2, i).Value = "" + newItem.DocReference;
                    _cells.GetCell(3, i).Value = "" + newItem.DocRefOther;
                    _cells.GetCell(4, i).Value = "" + newItem.DocumentNo;
                    _cells.GetCell(5, i).Value = "" + newItem.DocumentRef;
                    _cells.GetCell(6, i).Value = "" + newItem.DocPrefix;
                    _cells.GetCell(7, i).Value = "" + newItem.DocumentType;
                    _cells.GetCell(8, i).Value = "" + newItem.VerStatus;
                    _cells.GetCell(9, i).Value = "" + newItem.VerType;
                    _cells.GetCell(10, i).Value = "" + newItem.DocVerNo;
                    _cells.GetCell(11, i).Value = "" + newItem.DocumentName1;
                    _cells.GetCell(12, i).Value = "" + newItem.ElecRef;
                    _cells.GetCell(13, i).Value = "" + newItem.ElecType;

                    _cells.GetCell(ResultColumn03, i).Value = "ACTUALIZADO";
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateDocumentReference()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }

            _cells?.SetCursorDefault();
        }

        private void DeleteDocumentReference()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var district = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;

            var drOpContext = DocumentReferenceActions.GetDocRefOpContext(district, _frmAuth.EllipsePost, 100, Debugger.DebugWarnings);
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow03 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var refType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value));
                    var refNo = _cells.GetCell(2, i).Value;
                    var refOther = _cells.GetCell(3, i).Value;
                    var docNo = _cells.GetCell(4, i).Value;
                    var docRef = _cells.GetCell(5, i).Value;
                    var refPrefix = _cells.GetCell(6, i).Value;
                    var docType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value));
                    var versionType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value));
                    var versionStatus = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value));
                    var versionNo = _cells.GetCell(10, i).Value;
                    var docName = _cells.GetCell(11, i).Value;
                    var elecRef = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value));
                    var elecType = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value));

                    var item = new DocumentReferenceItem();
                    item.District = string.IsNullOrWhiteSpace(district) ? null : district;
                    item.DocRefType = string.IsNullOrWhiteSpace(refType) ? null : refType;
                    item.DocReference = string.IsNullOrWhiteSpace(refNo) ? null : refNo;
                    item.DocRefOther = string.IsNullOrWhiteSpace(refOther) ? null : refOther;
                    item.DocumentNo = string.IsNullOrWhiteSpace(docNo) ? null : docNo;
                    item.DocumentRef = string.IsNullOrWhiteSpace(docRef) ? null : docRef;
                    item.DocPrefix = string.IsNullOrWhiteSpace(refPrefix) ? null : refPrefix;
                    item.DocumentType = string.IsNullOrWhiteSpace(docType) ? null : docType;
                    item.VerType = string.IsNullOrWhiteSpace(versionType) ? null : versionType;
                    item.VerStatus = string.IsNullOrWhiteSpace(versionStatus) ? null : versionStatus;
                    item.DocVerNo = string.IsNullOrWhiteSpace(versionNo) ? null : versionNo;
                    item.DocumentName1 = string.IsNullOrWhiteSpace(docName) ? null : docName;
                    item.ElecRef = string.IsNullOrWhiteSpace(elecRef) ? null : elecRef;
                    item.ElecType = string.IsNullOrWhiteSpace(elecType) ? null : elecType;

                    var reply = DocumentReferenceActions.DeleteDocument(urlService, drOpContext, item);

                    if (reply.errors != null && reply.errors.Length > 0)
                    {
                        string error = "";
                        foreach (var err in reply.errors)
                        {
                            error = error + err.messageText + "\n";
                        }
                        throw new Exception(error);
                    }
                    var newItem = new DocumentReferenceItem(reply.documentReferenceDTO);

                    _cells.GetCell(1, i).Value = "" + newItem.DocRefType;
                    _cells.GetCell(2, i).Value = "" + newItem.DocReference;
                    _cells.GetCell(3, i).Value = "" + newItem.DocRefOther;
                    _cells.GetCell(4, i).Value = "" + newItem.DocumentNo;
                    _cells.GetCell(5, i).Value = "" + newItem.DocumentRef;
                    _cells.GetCell(6, i).Value = "" + newItem.DocPrefix;
                    _cells.GetCell(7, i).Value = "" + newItem.DocumentType;
                    _cells.GetCell(8, i).Value = "" + newItem.VerStatus;
                    _cells.GetCell(9, i).Value = "" + newItem.VerType;
                    _cells.GetCell(10, i).Value = "" + newItem.DocVerNo;
                    _cells.GetCell(11, i).Value = "" + newItem.DocumentName1;
                    _cells.GetCell(12, i).Value = "" + newItem.ElecRef;
                    _cells.GetCell(13, i).Value = "" + newItem.ElecType;

                    _cells.GetCell(ResultColumn03, i).Value = "ELIMINADO";
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:DeleteDocumentReference()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                }
            }

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
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        
    }
}
