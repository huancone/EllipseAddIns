using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Utilities;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Windows.Forms;
using LogsheetDatamodelLibrary;
using LogsheetDatamodelLibrary.Configuration;

namespace LogsheetDatamodelAdmin
{
    
    public partial class RibbonLsdm
    {
        private ExcelStyleCells _cells;
        private AuthenticationForm _frmAuth;
        private Excel.Application _excelApp;

        //Datasheet
        private const string SheetNameDatasheet = "Datasheet";
        private const string TableNameDatasheet = "DatasheetTable";
        private const int TitleRowDatasheet = 9;
        private const int ResultColumnDatasheet = 4;

        //Model
        private const string SheetNameModel = "Model";
        private const string TableNameModel = "ModelTable";
        private const int TitleRowModel = 7;
        private const int ResultColumnModel = 8;
        //Attributes
        private const string SheetNameAttribute = "Attributes";
        private const string TableNameAttribute = "AttributesTable";
        private const int TitleRowAttribute = 7;
        private const int ResultColumnAttribute = 14;

        //Measure
        private const string SheetNameMeasure = "Measure";
        private const string TableNameMeasure = "MeasureTable";
        private const int TitleRowMeasure = 7;
        private const int ResultColumnMeasure = 8;

        //MeasureType
        private const string SheetNameMeasureType = "MeasureType";
        private const string TableNameMeasureType = "MeasureTypeTable";
        private const int TitleRowMeasureType = 7;
        private const int ResultColumnMeasureType = 3;

        //Validation Items
        private const string SheetNameValidItems = "ValidItems";
        private const string TableNameValidItems = "ValidItemsTable";
        private const int TitleRowValidItems = 7;
        private const int ResultColumnValidItems = 8;
        //Validation Sources
        private const string SheetNameValidSources = "ValidSources";
        private const string TableNameValidSources = "ValidSourcesTable";
        private const int TitleRowValidSources = 7;
        private const int ResultColumnValidSources = 8;

        private const string ValidationSheetName = "ValidationSheet";
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            var settings = new Settings();

            _frmAuth = new AuthenticationForm();
            _excelApp = Globals.ThisAddIn.Application;

            var defaultConfig = new SharedClassLibrary.Configuration.Options();
            //defaultConfig.SetOption("OptionName1", "OptionValue1");
            //defaultConfig.SetOption("OptionName2", "OptionValue2");
            //defaultConfig.SetOption("OptionName3", "OptionValue3");

            var options = settings.GetOptionsSettings(defaultConfig);

            //Setting of Configuration Options from Config File (or default)
            //var optionItem1Value = MyUtilities.IsTrue(options.GetOptionValue("OptionName1"));
            //var optionItem1Value = options.GetOptionValue("OptionName2");
            //var optionItem1Value = options.GetOptionValue("OptionName3");

            //optionItem1.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.UpdateOptionsSettings(options);
            
            LsdmConfig.Settings = settings;
            LsdmConfig.DataSource = new DataSource();
            LsdmConfig.DataSource.SetDBSettings("XEPDB1", "lsdmuser", "12345");
        }

        private void btnFormat_Click(object sender, RibbonControlEventArgs e)
        {
            FormatMethod();
            if (!_cells.IsDecimalDotSeparator())
                MessageBox.Show(Resources.Warning_DecimalSeparatorWarning, Resources.Warning_WarningUppercase);
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            const string developerName1 = "Héctor Hernandez <hernandezrhectorj@gmail.com>";
            const string developerName2 = "Hugo Mendoza <huancone@gmail.com>";

            new AboutBox(developerName1, developerName2).ShowDialog();
        }

        private void btnStop_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show($@"{Resources.Error_ThreadStopped} . {ex.Message}");
            }
        }

        private void FormatMethod()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                while (_excelApp.ActiveWorkbook.Sheets.Count < 7)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                #region Datasheet

                var validationColumn = 1;

                var titleRow = TitleRowDatasheet;
                var resultColumn = ResultColumnDatasheet;
                var tableName = TableNameDatasheet;
                var sheetName = SheetNameDatasheet;
                var currentSheet = 1;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.DatasheetUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                //SEARCH
                _cells.GetCell("A4").Value = LsdmResource.Datasheet_ModelId;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                _cells.GetCell("A5").Value = Resources.Search_HideInactive;
                _cells.GetCell("A5").Style = StyleConstants.Option;
                _cells.GetCell("B5").Value = Resources.Yes_Camelcase;
                _cells.GetCell("B5").Style = StyleConstants.Select;

                _cells.GetCell("A6").Value = LsdmResource.Search_StartDate;
                _cells.GetCell("A6").Style = StyleConstants.Option;
                _cells.GetCell("B6").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month)  + "01";
                _cells.GetCell("B6").Style = StyleConstants.Select;
                _cells.GetCell("B6").AddComment("YYYYMMDD");
                _cells.GetCell("A7").Value = LsdmResource.Search_FinishDate;
                _cells.GetCell("A7").Style = StyleConstants.Option;
                _cells.GetCell("B7").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("B7").Style = StyleConstants.Select;
                _cells.GetCell("B7").AddComment("YYYYMMDD");

                var activeList = new List<string> { Resources.Yes_Camelcase, Resources.No_Camelcase };

                _cells.SetValidationList(_cells.GetCell("B5"), activeList);
                //ATTRIBUTES

                _cells.GetCell(1, titleRow).Value = LsdmResource.Datasheet_Date;
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = LsdmResource.Datasheet_Shift;
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = LsdmResource.Datasheet_SequenceId;
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                var shiftList = new List<string>();//TO DO

                _cells.SetValidationList(_cells.GetCell(2, titleRow + 1), shiftList, ValidationSheetName, validationColumn, false);
                validationColumn++;
                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region MODEL ATTRIBUTES
                titleRow = TitleRowAttribute;
                resultColumn = ResultColumnAttribute;
                tableName = TableNameAttribute;
                sheetName = SheetNameAttribute;
                currentSheet++;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.ModelUppercase + " & " + LsdmResource.AttributesUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                //SEARCH
                _cells.GetCell("A4").Value = LsdmResource.Attribute_ModelId;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                _cells.GetCell("A5").Value = Resources.Search_HideInactive;
                _cells.GetCell("A5").Style = StyleConstants.Option;
                _cells.GetCell("B5").Value = Resources.Yes_Camelcase;
                _cells.GetCell("B5").Style = StyleConstants.Select;

                activeList = new List<string> {Resources.Yes_Camelcase, Resources.No_Camelcase};

                _cells.SetValidationList(_cells.GetCell("B5"), activeList);
                //ATTRIBUTES

                _cells.GetCell(1, titleRow).Value = LsdmResource.Attribute_ModelId;
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = LsdmResource.Attribute_Id;
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = LsdmResource.Attribute_Description;
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(4, titleRow).Value = LsdmResource.Attribute_DataType;
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, titleRow).Value = LsdmResource.Attribute_SheetIndex;
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(6, titleRow).Value = LsdmResource.Attribute_MaxLength;
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(7, titleRow).Value = LsdmResource.Attribute_MaxPrecision;
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(8, titleRow).Value = LsdmResource.Attribute_MaxScale;
                _cells.GetCell(8, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(9, titleRow).Value = LsdmResource.Attribute_AllowNull;
                _cells.GetCell(9, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(10, titleRow).Value = LsdmResource.Attribute_DefaultValue;
                _cells.GetCell(10, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(11, titleRow).Value = LsdmResource.Attribute_Status;
                _cells.GetCell(11, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(12, titleRow).Value = LsdmResource.Attribute_MeasureId;
                _cells.GetCell(12, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(13, titleRow).Value = LsdmResource.Attribute_ValidItemId;
                _cells.GetCell(13, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                var dataTypeList = DataTypes.GetList();

                var measureList = Measure.Read();
                var measureValidList = measureList.Select(m => m.Id + " - " + m.Code + " " + m.Name).ToList();

                var validationList = ValidationItem.Read();
                var validItemList = validationList.Select(v => v.Id + " - " + v.Description).ToList();

                _cells.SetValidationList(_cells.GetCell(4, titleRow + 1), dataTypeList, ValidationSheetName, validationColumn++);
                _cells.SetValidationList(_cells.GetCell(12, titleRow + 1), measureValidList, ValidationSheetName, validationColumn++, false);
                _cells.SetValidationList(_cells.GetCell(13, titleRow + 1), validItemList, ValidationSheetName, validationColumn++, false);

                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region DATAMODEL
                titleRow = TitleRowModel;
                resultColumn = ResultColumnModel;
                tableName = TableNameModel;
                sheetName = SheetNameModel;
                currentSheet++;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.ModelUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                //SEARCH
                _cells.GetCell("A4").Value = LsdmResource.Model_Id;
                _cells.GetCell("A5").Value = LsdmResource.Search_Keyword;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("A5").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                _cells.GetCell("B5").Style = StyleConstants.Select;

                //MODEL
                _cells.GetCell(1, titleRow).Value = LsdmResource.Model_Id;
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(2, titleRow).Value = LsdmResource.Model_Description;
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = LsdmResource.Model_Status;
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(4, titleRow).Value = LsdmResource.Model_CreationDate;
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(5, titleRow).Value = LsdmResource.Model_CreationUser;
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(6, titleRow).Value = LsdmResource.Model_LastModDate;
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(7, titleRow).Value = LsdmResource.Model_LastModUser;
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleInformation);


                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region MEASURE
                titleRow = TitleRowMeasure;
                resultColumn = ResultColumnMeasure;
                tableName = TableNameMeasure;
                sheetName = SheetNameMeasure;
                currentSheet++;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.MeasureUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetCell("A4").Value = LsdmResource.Measure_Id;
                _cells.GetCell("A5").Value = LsdmResource.Measure_Code;
                _cells.GetCell("C4").Value = LsdmResource.Search_Keyword;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("A5").Style = StyleConstants.Option;
                _cells.GetCell("C4").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                _cells.GetCell("B5").Style = StyleConstants.Select;
                _cells.GetCell("D4").Style = StyleConstants.Select;

                //
                _cells.GetCell(1, titleRow).Value = LsdmResource.Measure_Id;
                _cells.GetCell(1, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(2, titleRow).Value = LsdmResource.Measure_Code;
                _cells.GetCell(2, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(3, titleRow).Value = LsdmResource.Measure_Name;
                _cells.GetCell(3, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(4, titleRow).Value = LsdmResource.Measure_Description;
                _cells.GetCell(4, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(5, titleRow).Value = LsdmResource.Measure_Units;
                _cells.GetCell(5, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(6, titleRow).Value = LsdmResource.Measure_Status;
                _cells.GetCell(6, titleRow).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(7, titleRow).Value = LsdmResource.Measure_TypeId;
                _cells.GetCell(7, titleRow).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                var measureTypeList = MeasureType.Read();
                var measureTypeValidList = measureTypeList.Select(mt => mt.Id + " - " + mt.Description).ToList();
                _cells.SetValidationList(_cells.GetCell(7, titleRow + 1), measureTypeValidList, ValidationSheetName, validationColumn++, false);

                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region MEASURE_TYPE

                titleRow = TitleRowMeasureType;
                resultColumn = ResultColumnMeasureType;
                tableName = TableNameMeasureType;
                sheetName = SheetNameMeasureType;
                currentSheet++;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.MeasureTypeUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetCell("A4").Value = LsdmResource.MeasureType_Id;
                _cells.GetCell("A5").Value = LsdmResource.MeasureType_Description;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("A5").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                _cells.GetCell("B5").Style = StyleConstants.Select;
                //
                _cells.GetCell(1, titleRow).Value = LsdmResource.MeasureType_Id;
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(2, titleRow).Value = LsdmResource.MeasureType_Description;
                _cells.GetCell(2, titleRow).Style = StyleConstants.TitleRequired;

                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region VALIDATION ITEMS

                titleRow = TitleRowValidItems;
                resultColumn = ResultColumnValidItems;
                tableName = TableNameValidItems;
                sheetName = SheetNameValidItems;
                currentSheet++;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.ValidationItemUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetCell("A4").Value = LsdmResource.Search_Keyword;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                //
                _cells.GetCell(1, titleRow).Value = LsdmResource.ValidationItem_SourceName;
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, titleRow).Value = LsdmResource.ValidationItem_Id;
                _cells.GetCell(2, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, titleRow).Value = LsdmResource.ValidationItem_Description;
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(4, titleRow).Value = LsdmResource.ValidationItem_SourceTable;
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, titleRow).Value = LsdmResource.ValidationItem_SourceColumn;
                _cells.GetCell(5, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, titleRow).Value = LsdmResource.ValidationItem_Sortable;
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(7, titleRow).Value = LsdmResource.ValidationItem_DistincFilter;
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleRequired;

                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region VALIDATION SOURCES

                titleRow = TitleRowValidSources;
                resultColumn = ResultColumnValidSources;
                tableName = TableNameValidSources;
                sheetName = SheetNameValidSources;
                currentSheet++;

                _excelApp.ActiveWorkbook.Sheets.get_Item(currentSheet).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = sheetName;

                _cells.GetCell("A1").Value = LsdmResource.LogsheetDatamodelerUppercase;
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = LsdmResource.ValidationSourcesUppercase;
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = Resources.Validation_MandatoryUppercase;
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = Resources.Validation_OptionalUppercase;
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = Resources.Validation_InformationUppercase;
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K4").Value = Resources.Validation_ActionToDoUppercase;
                _cells.GetCell("K4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("K5").Value = Resources.Validation_AdditionalRequiredUppercase;
                _cells.GetCell("K5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetCell("A4").Value = LsdmResource.Search_Keyword;
                _cells.GetCell("A4").Style = StyleConstants.Option;
                _cells.GetCell("B4").Style = StyleConstants.Select;
                //
                _cells.GetCell(1, titleRow).Value = LsdmResource.ValidationSource_Name;
                _cells.GetCell(1, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(2, titleRow).Value = LsdmResource.ValidationSource_DbName;
                _cells.GetCell(2, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(3, titleRow).Value = LsdmResource.ValidationSource_DbUser;
                _cells.GetCell(3, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(4, titleRow).Value = LsdmResource.ValidationSource_DbPassword;
                _cells.GetCell(4, titleRow).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, titleRow).Value = LsdmResource.ValidationSource_DbReference;
                _cells.GetCell(5, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, titleRow).Value = LsdmResource.ValidationSource_DbLink;
                _cells.GetCell(6, titleRow).Style = StyleConstants.TitleOptional;
                _cells.GetCell(7, titleRow).Value = LsdmResource.ValidationSource_EncodedPasswordType;
                _cells.GetCell(7, titleRow).Style = StyleConstants.TitleOptional;

                var encryptList = ValidationSource.EncryptionTypeValues.GetList();
                _cells.SetValidationList(_cells.GetCell(7, titleRow + 1), encryptList, ValidationSheetName, validationColumn+1);

                _cells.GetCell(resultColumn, titleRow).Value = Resources.Title_ResultCamelcase;
                _cells.GetCell(resultColumn, titleRow).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, titleRow + 1, resultColumn, titleRow + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, titleRow, resultColumn, titleRow + 1), tableName);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatMehod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show($@"{Resources.Error_SheetHeaderError} . {ex.Message}");
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }


        private void btnAttributeSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameAttribute)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(AttributeSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:AttributeSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnAttributeSearchEach_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet) _excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameAttribute)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(AttributeSearchEachMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:AttributeSearchEachMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);

            }
        }

        private void btnAttributeUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameAttribute)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(AttributeUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:AttributeUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnAttributeDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameAttribute)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(AttributeDeleteMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:AttributeDeleteMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }
        private void btnModelSearch_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameModel)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ModelSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ModelSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }

        }

        private void btnModelSearchEach_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameModel)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ModelSearchEachMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ModelSearchEachMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnModelUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameModel)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ModelUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ModelUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnModelDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameModel)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ModelDeleteMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ModelDeleteMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }


        private void btnMeasureSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasure)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }
        private void btnMeasureSearchEach_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasure)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureSearchEachMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }
        private void btnMeasureUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasure)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnMeasureDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasure)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureDeleteMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureDeleteMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }


        private void btnMeasureTypeSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasureType)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureTypeSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureTypeSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }
        private void btnMeasureTypeSearchEach_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasureType)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureTypeSearchEachMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureTypeSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }
        private void btnMeasureTypeUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasureType)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureTypeUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureTypeUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnMeasureTypeDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameMeasureType)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(MeasureTypeDeleteMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:MeasureTypeDeleteMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidItemsSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidItems)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationItemsSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationItemsSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidItemsSearchEach_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidItems)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationItemsSearchEachMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationItemsSearchEachMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidItemsUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidItems)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationItemsUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationItemsUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidItemsDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidItems)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationItemsDeleteMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationItemsDeleteMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidSourcesSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidSources)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationSourcesSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationSourcesSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidSourcesSearchEach_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidSources)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationSourcesSearchEachMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationSourcesSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidSourcesUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidSources)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationSourcesUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationSourcesUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnValidSourcesDelete_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameValidSources)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(ValidationSourcesDeleteMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:ValidationSourcesDeleteMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnDatasheetSearch_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameDatasheet)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(DatasheetSearchMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DatasheetSearchMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnDatasheetUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (((Worksheet)_excelApp.ActiveWorkbook.ActiveSheet).Name == SheetNameDatasheet)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;

                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(DatasheetUpdateMethod);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(Resources.Error_ExcelSheetFormatError);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:DatasheetUpdateMethod()", "\n\r" + Resources.Debugging_Message + ":" + ex.Message + "\n\r" + Resources.Debugging_Source + ":" + ex.Source + "\n\r" + Resources.Debugging_StackTrace + ":" + ex.StackTrace);
                MessageBox.Show(Resources.Error_ErrorFound + @": " + ex.Message);
            }
        }

        private void btnDatasheetDelete_Click(object sender, RibbonControlEventArgs e)
        {

        }

        
    }
}
