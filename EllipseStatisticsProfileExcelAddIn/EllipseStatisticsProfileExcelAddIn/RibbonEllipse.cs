using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = SharedClassLibrary.Ellipse.ScreenService;
using Debugger = SharedClassLibrary.Utilities.Debugger;

namespace EllipseStatisticsProfileExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        
        private readonly string _sheetName01 = "Statistics Profile";

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
        private void btnFormatProfile_Click(object sender, RibbonControlEventArgs e)
        {
            FormatProfile();
        }

        private void FormatProfile()
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

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = _sheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STATISTICS PROFILES - ELLIPSE";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("A4:AT4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("AU4").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("C4").Value = "Equipment";
                _cells.GetCell("D4").Value = "Fuel Type";
                _cells.GetCell("E4").Value = "Fuel Capacity";
                _cells.GetCell("F4").Value = "Primary Statistic";
                _cells.GetCell("AU4").Value = "Result";

                var optionList = new List<string>();
                _cells.GetCell("A4").Value = "Options";
                optionList.Add("3. Maintain Operating Statistics Profile");
                optionList.Add("4. Maintain General Equipment Profile");
                optionList.Add("Delete Profile");
                _cells.SetValidationList(_cells.GetCell("A5:A100"), optionList);

                _cells.GetCell("B4").Value = "Profile Type";
                
                optionList = new List<string> { "EGI", "Equipment" };
                _cells.SetValidationList(_cells.GetCell("B5:B100"), optionList);

                const int startColumn = 7;
                for (var i = 0; i < 20; i++)
                {
                    _cells.GetRange(i * 2 + startColumn, 3, i * 2 + startColumn + 1, 3).Style =
                        _cells.GetStyle(StyleConstants.Option);
                    _cells.MergeCells(i * 2 + startColumn, 3, i * 2 + startColumn + 1, 3);
                    _cells.GetCell(i * 2 + startColumn, 3).Value = i + 1;
                    _cells.GetCell(i * 2 + startColumn, 4).Value = "Stat " + (i + 1);
                    _cells.GetCell(i * 2 + startColumn + 1, 4).Value = "D/I";
                }

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                _cells.GetCell("AU5").Value = error.Message;
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void btnExecuteProfile_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteProfile();
        }

        private void ExecuteProfile()
        {
            var statType = new List<string>();
            var statEntry = new List<string>();
            var arrayFields = new ArrayScreenNameValue();

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

            if (_frmAuth.ShowDialog() != DialogResult.OK) return;
            var opSheet = new Screen.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var proxySheet = new Screen.ScreenService();
            var requestSheet = new Screen.ScreenSubmitRequestDTO();
            var replySheet = new Screen.ScreenDTO();

            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";

            var currentRow = 5;

            string option1I = _cells.GetEmptyIfNull(_cells.GetCell("A" + currentRow).Value);

            try
            {
                while (!string.IsNullOrEmpty(option1I))
                {
                    string profile;
                    string profileType;
                    switch (option1I)
                    {
                        case "3. Maintain Operating Statistics Profile":
                            string prStatType1I;
                            MaintainStatisticsProfile(ref option1I, statType, statEntry, ref arrayFields, opSheet,
                                proxySheet, requestSheet, ref replySheet, currentRow, out profile, out profileType,
                                out prStatType1I);
                            break;
                        case "4. Maintain General Equipment Profile":
                            string fuelOilType2I;
                            string fuelCapacity2I;
                            ManintainGeneralProfile(ref option1I, ref arrayFields, opSheet, proxySheet, requestSheet,
                                ref replySheet, currentRow, out profile, out profileType, out fuelOilType2I,
                                out fuelCapacity2I);
                            break;
                        case "Delete Profile":
                            DeleteProfile(ref option1I, ref arrayFields, opSheet, proxySheet, requestSheet,
                                ref replySheet, currentRow, out profile, out profileType);
                            break;
                    }
                    currentRow += 1;
                    option1I = "" + _cells.GetCell("A" + currentRow).Value;
                }
            }
            catch (Exception error)
            {
                _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                _cells.GetCell("A" + currentRow).AddComment("" + error.Message);
            }
        }

        private void MaintainStatisticsProfile(ref string option1I, List<string> statType, List<string> statEntry,
            ref ArrayScreenNameValue arrayFields, Screen.OperationContext opSheet, Screen.ScreenService proxySheet,
            Screen.ScreenSubmitRequestDTO requestSheet, ref Screen.ScreenDTO replySheet, int currentRow,
            out string profile, out string profileType, out string prStatType1I)
        {
            option1I = option1I.Substring(0, 1);
            profileType = "" + _cells.GetCell("B" + currentRow).Value;
            profile = "" + _cells.GetCell("C" + currentRow).Value;
            prStatType1I = "" + _cells.GetCell("F" + currentRow).Value;

            statType = new List<string>();
            statEntry = new List<string>();

            const int startColumn = 7;
            for (var i = 0; i < 20; i++)
            {
                statType.Add(_cells.GetEmptyIfNull(_cells.GetCell(i * 2 + startColumn, currentRow).Value));
                statEntry.Add(_cells.GetEmptyIfNull(_cells.GetCell(i * 2 + startColumn + 1, currentRow).Value));
            }

            _eFunctions.RevertOperation(opSheet, proxySheet);
            replySheet = proxySheet.executeScreen(opSheet, "MSO615");

            if (replySheet.mapName != "MSM615A" || _excelApp.ActiveWorkbook.ActiveSheet.Name != _sheetName01) return;
            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", option1I);
            arrayFields.Add(profileType == "EGI" ? "GROUP_ID11I" : "PLANT_NO11I", profile);

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";
            replySheet = proxySheet.submit(opSheet, requestSheet);

            if (_eFunctions.CheckReplyWarning(replySheet))
                replySheet = proxySheet.submit(opSheet, requestSheet);

            if (replySheet != null && replySheet.mapName == "MSM617A")
            {
                arrayFields = new ArrayScreenNameValue();

                arrayFields.Add("EQUIP_REF1I", profile);
                arrayFields.Add("PR_STAT_TYPE1I", prStatType1I);

                double index = 0;
                double firstIndex;
                double secondIndex;
                foreach (var item in statType)
                {
                    firstIndex = (index % 2) + 1;
                    secondIndex = Math.Ceiling((index + 1) / 2);
                    arrayFields.Add("STAT_TYPE" + firstIndex + "1I" + secondIndex, item);
                    index++;
                }

                index = 0;
                foreach (var item in statEntry)
                {
                    firstIndex = (index % 2) + 1;
                    secondIndex = Math.Ceiling((index + 1) / 2);
                    arrayFields.Add("STAT_ENTRY" + firstIndex + "1I" + secondIndex, item);
                    index++;
                }

                requestSheet.screenFields = arrayFields.ToArray();
                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                while (replySheet != null && replySheet.mapName == "MSM617A" &&
                       (replySheet.functionKeys.Contains("XMIT-Confirm") ||
                        replySheet.functionKeys.Contains("XMIT-Validate") ||
                        _eFunctions.CheckReplyWarning(replySheet)))
                {
                    replySheet = proxySheet.submit(opSheet, requestSheet);
                }

                if (replySheet != null && replySheet.mapName == "MSM615A")
                {
                    _cells.GetCell("A" + currentRow + ":" + "C" + currentRow).Style =
                        _cells.GetStyle(StyleConstants.Success);
                }
                else if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style =
                        _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell("A" + currentRow).AddComment(_cells.GetEmptyIfNull(replySheet.message));
                }
                else
                {
                    _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style =
                        _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell("A" + currentRow).AddComment("NPI");
                }
            }
            else if (_eFunctions.CheckReplyError(replySheet))
            {
                _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                _cells.GetCell("A" + currentRow).AddComment(_cells.GetEmptyIfNull(replySheet.message));
            }
            else
            {
                _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                _cells.GetCell("A" + currentRow).AddComment("NPI");
            }
        }

        private void ManintainGeneralProfile(ref string option1I, ref ArrayScreenNameValue arrayFields,
            Screen.OperationContext opSheet, Screen.ScreenService proxySheet, Screen.ScreenSubmitRequestDTO requestSheet,
            ref Screen.ScreenDTO replySheet, int currentRow, out string profile, out string profileType,
            out string fuelOilType2I, out string fuelCapacity2I)
        {
            option1I = option1I.Substring(0, 1);
            profileType = "" + _cells.GetCell("B" + currentRow).Value;
            profile = "" + _cells.GetCell("C" + currentRow).Value;
            fuelOilType2I = "" + _cells.GetCell("D" + currentRow).Value;
            fuelCapacity2I = "" + _cells.GetCell("E" + currentRow).Value;

            _eFunctions.RevertOperation(opSheet, proxySheet);
            replySheet = proxySheet.executeScreen(opSheet, "MSO615");

            if (replySheet.mapName == "MSM615A" && _excelApp.ActiveWorkbook.ActiveSheet.Name == _sheetName01)
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("OPTION1I", option1I);

                if (profileType == "EGI")
                {
                    arrayFields.Add("GROUP_ID11I", profile);
                }
                else
                {
                    arrayFields.Add("PLANT_NO11I", profile);
                }

                requestSheet.screenFields = arrayFields.ToArray();
                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);

                if (_eFunctions.CheckReplyWarning(replySheet))
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                if (replySheet != null && replySheet.mapName == "MSM617B" && !_eFunctions.CheckReplyError(replySheet))
                {
                    arrayFields = new ArrayScreenNameValue();

                    arrayFields.Add("PLANT_NO2I", profile);
                    arrayFields.Add("FUEL_OIL_TYPE2I", _cells.GetNullIfTrimmedEmpty(fuelOilType2I) == null ? replySheet.screenFields[9].value : fuelOilType2I);
                    arrayFields.Add("FUEL_CAPACITY2I",   _cells.GetNullIfTrimmedEmpty(fuelCapacity2I) == null ? replySheet.screenFields[1].value : fuelCapacity2I) ;
                    arrayFields.Add("VAL_PROF_FLG2I", "N");

                    requestSheet.screenFields = arrayFields.ToArray();
                    requestSheet.screenKey = "1";
                    replySheet = proxySheet.submit(opSheet, requestSheet);

                    if (replySheet != null && replySheet.mapName == "MSM617B" &&
                        (replySheet.functionKeys.Contains("XMIT-Confirm") || _eFunctions.CheckReplyWarning(replySheet)))
                    {
                        replySheet = proxySheet.submit(opSheet, requestSheet);

                        if (replySheet != null && replySheet.mapName == "MSM615A" &&
                            !_eFunctions.CheckReplyWarning(replySheet) && !_eFunctions.CheckReplyError(replySheet))
                        {
                            _cells.GetCell("A" + currentRow + ":" + "C" + currentRow).Style =
                                _cells.GetStyle(StyleConstants.Success);
                        }
                    }
                    else if (_eFunctions.CheckReplyError(replySheet))
                    {
                        _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style =
                            _cells.GetStyle(StyleConstants.Error);
                        _cells.GetCell("A" + currentRow).AddComment(_cells.GetEmptyIfNull(replySheet.message));
                    }
                    else
                    {
                        _cells.GetCell("A" + currentRow + ":" + "C" + currentRow).Style =
                            _cells.GetStyle(StyleConstants.Success);
                    }
                }
                else if (_eFunctions.CheckReplyError(replySheet))
                {
                    _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell("A" + currentRow).AddComment(_cells.GetEmptyIfNull(replySheet.message));
                }
                else
                {
                    _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style = _cells.GetStyle(StyleConstants.Error);
                    _cells.GetCell("A" + currentRow).AddComment("NPI");
                }
            }
        }

        private void DeleteProfile(ref string option1I, ref ArrayScreenNameValue arrayFields,
            Screen.OperationContext opSheet, Screen.ScreenService proxySheet, Screen.ScreenSubmitRequestDTO requestSheet, ref Screen.ScreenDTO replySheet, int currentRow, out string profile, out string profileType)
        {
            option1I = "3";
            profileType = "" + _cells.GetCell("B" + currentRow).Value;
            profile = "" + _cells.GetCell("C" + currentRow).Value;

            _eFunctions.RevertOperation(opSheet, proxySheet);
            replySheet = proxySheet.executeScreen(opSheet, "MSO615");

            if (replySheet.mapName != "MSM615A" || _excelApp.ActiveWorkbook.ActiveSheet.Name != _sheetName01) return;
            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", option1I);

            arrayFields.Add(profileType == "EGI" ? "GROUP_ID11I" : "PLANT_NO11I", profile);

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "1";
            replySheet = proxySheet.submit(opSheet, requestSheet);

            if (_eFunctions.CheckReplyWarning(replySheet))
                replySheet = proxySheet.submit(opSheet, requestSheet);

            if (replySheet == null || replySheet.mapName != "MSM617A" || _eFunctions.CheckReplyError(replySheet))
                return;
            arrayFields = new ArrayScreenNameValue();

            arrayFields.Add("EQUIP_REF1I", profile);

            requestSheet.screenFields = arrayFields.ToArray();
            requestSheet.screenKey = "9";
            replySheet = proxySheet.submit(opSheet, requestSheet);

            if (_eFunctions.CheckReplyWarning(replySheet))
                replySheet = proxySheet.submit(opSheet, requestSheet);

            while (replySheet.mapName == "MSM617A")
            {
                arrayFields = new ArrayScreenNameValue();
                arrayFields.Add("CONF1I", "Y");

                requestSheet.screenFields = arrayFields.ToArray();
                requestSheet.screenKey = "1";
                replySheet = proxySheet.submit(opSheet, requestSheet);
            }

            if (replySheet != null && replySheet.mapName == "MSM615A" && !_eFunctions.CheckReplyError(replySheet))
            {
                _cells.GetCell("A" + currentRow + ":" + "C" + currentRow).Style =
                    _cells.GetStyle(StyleConstants.Success);
            }
            else if (_eFunctions.CheckReplyError(replySheet))
            {
                _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style =
                    _cells.GetStyle(StyleConstants.Error);
                _cells.GetCell("A" + currentRow).AddComment(_cells.GetEmptyIfNull(replySheet.message));
            }
            else
            {
                _cells.GetCell("A" + currentRow + ":" + "B" + currentRow).Style =
                    _cells.GetStyle(StyleConstants.Error);
                _cells.GetCell("A" + currentRow).AddComment("NPI");
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}