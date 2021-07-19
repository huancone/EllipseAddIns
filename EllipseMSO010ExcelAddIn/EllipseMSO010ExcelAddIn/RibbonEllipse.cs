using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Classes;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using Screen = SharedClassLibrary.Ellipse.ScreenService; //si es screen service
using System.Web.Services.Ellipse;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;

namespace EllipseMSO010ExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const string SheetName01 = "MSO010 Codes";
        private const int TitleRow01 = 9;
        private const int ResultColumn01 = 7;
        private const string TableName01 = "CodesTable";
        private const string ValidationSheetName = "ValidationSheet";
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
        private void btnFormatCesantias_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }
        private void btnCreate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(CreateCodeList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreadeCodeList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
        private void btnReview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ReviewCodesList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewCodesList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void btnModify_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(ModifyCodeList);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ModifyCodeList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
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
                _excelApp.ActiveWorkbook.Worksheets.Add();//hoja 4
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "MSO010 TABLE CODE ENTRIES - ELLIPSE";
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

                var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes().Select(g => g.Value).ToList();
                var typeCriteriaList = SearchTypeCriteriaType.GetSearchTypeCriteriaTypes().Select(g => g.Value).ToList();
                var statusList = StatusCode.GetStatusList().Select(g => g.Value).ToList();
                _cells.GetCell("A3").Value = SearchFieldCriteriaType.Type.Value;
                _cells.GetCell("B3").Value = SearchTypeCriteriaType.EqualsTo.Value;
                _cells.SetValidationList(_cells.GetCell("A3"), searchCriteriaList, ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B3"), typeCriteriaList, ValidationSheetName, 3);
                _cells.GetCell("A4").Value = SearchFieldCriteriaType.Code.Value;
                _cells.GetCell("B4").Value = SearchTypeCriteriaType.EqualsTo.Value;
                _cells.SetValidationList(_cells.GetCell("A4"), ValidationSheetName, 2);
                _cells.SetValidationList(_cells.GetCell("B4"), ValidationSheetName, 3);

                _cells.GetCell("A5").Value = "STATUS";
                _cells.GetCell("B5").Value = StatusCode.Active.Value;
                _cells.SetValidationList(_cells.GetCell("B5"), statusList, ValidationSheetName, 4);
                
                _cells.GetRange("A3", "B5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("B5").Style = _cells.GetStyle(StyleConstants.Select);


                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(02, TitleRow01).Style = StyleConstants.TitleInformation;
                //GENERAL

                _cells.GetCell(01, TitleRow01).Value = "TYPE";
                _cells.GetCell(02, TitleRow01).Value = "TYPE DESC";
                _cells.GetCell(03, TitleRow01).Value = "CODE";
                _cells.GetCell(04, TitleRow01).Value = "STATUS";
                _cells.GetCell(05, TitleRow01).Value = "DESCRIPTION";
                _cells.GetCell(06, TitleRow01).Value = "ASSOC_REC";
                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;

                var actStatusList = new List<string> {"Y - Active", "N - Inactive"};
                _cells.SetValidationList(_cells.GetCell(04, TitleRow01 + 1), typeCriteriaList, ValidationSheetName, 4);
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatQuality()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja" + "\n" + ex.Message);
            }
        }
       
        public void ReviewCodesList()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableName01);

            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var searchCriteriaList = SearchFieldCriteriaType.GetSearchFieldCriteriaTypes();
            var typeCriteriaList = SearchTypeCriteriaType.GetSearchTypeCriteriaTypes();
            var statusCriteriaList = StatusCode.GetStatusList();

            //Obtengo los valores de las opciones de búsqueda
            var searchCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("A3").Value);
            var typeCriteriaKey1Text = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value);
            var searchCriteriaValue1 = _cells.GetEmptyIfNull(_cells.GetCell("C3").Value).ToUpper();
            var searchCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("A4").Value);
            var typeCriteriaKey2Text = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value);
            var searchCriteriaValue2 = _cells.GetEmptyIfNull(_cells.GetCell("C4").Value).ToUpper(); 
            var statusKeyText = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value);

            //Convierto los nombres de las opciones a llaves
            var searchCriteriaKey1 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey1Text)).Key;
            var typeCriteriaKey1 = typeCriteriaList.FirstOrDefault(v => v.Value.Equals(typeCriteriaKey1Text)).Key;
            var searchCriteriaKey2 = searchCriteriaList.FirstOrDefault(v => v.Value.Equals(searchCriteriaKey2Text)).Key;
            var typeCriteriaKey2 = typeCriteriaList.FirstOrDefault(v => v.Value.Equals(typeCriteriaKey2Text)).Key;
            var statusKey = statusCriteriaList.FirstOrDefault(v => v.Value.Equals(statusKeyText)).Key;

            var listwo = GetItemList(_eFunctions, searchCriteriaKey1, typeCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, typeCriteriaKey2, searchCriteriaValue2, statusKey);
            var i = TitleRow01 + 1;
            foreach (var item in listwo)
            {
                try
                {
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "'" + item.Type;
                    _cells.GetCell(2, i).Value = "" + item.TypeDescription;
                    _cells.GetCell(3, i).Value = "'" + item.Code;
                    _cells.GetCell(4, i).Value = "'" + item.ActiveStatus;
                    _cells.GetCell(5, i).Value = "" + item.Description;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewCodesList()", ex.Message);
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

        public void CreateCodeList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
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

            while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)) || !string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value2)))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var item = new ItemCode();
                    //GENERAL
                    item.Type = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value).ToUpper();
                    item.TypeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    item.Code = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value).ToUpper();
                    item.ActiveStatus = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value));
                    item.Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    item.AssocRec = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);

                    item.ActiveStatus = string.IsNullOrWhiteSpace(item.ActiveStatus) ? "Y" : item.ActiveStatus;

                    CreateCodeRegister(opSheet, proxySheet, item);

                    _cells.GetCell(ResultColumn01, i).Value = "REGISTRO CREADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateCodeList()", ex.Message);
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

        public void ModifyCodeList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
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

            while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)) || !string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(2, i).Value2)))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var item = new ItemCode();
                    //GENERAL
                    item.Type = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value).ToUpper();
                    item.TypeDescription = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                    item.Code = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value).ToUpper();
                    item.ActiveStatus = MyUtilities.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value));
                    item.Description = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value);
                    item.AssocRec = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value);

                    item.ActiveStatus = string.IsNullOrWhiteSpace(item.ActiveStatus) ? "Y" : item.ActiveStatus;

                    ModifyCodeRegister(opSheet, proxySheet, item);

                    _cells.GetCell(ResultColumn01, i).Value = "REGISTRO CREADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateCodeList()", ex.Message);
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

        public void CreateCodeRegister(Screen.OperationContext opContext, Screen.ScreenService proxySheet, ItemCode item)
        {
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO010");
            //Validamos el ingreso
            if (reply.mapName != "MSM010A")
                throw new Exception("NO SE PUEDE INGRESAR AL PROGRAMA MSO010");
            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("TABLE_TYPE1I", item.Type);
            arrayFields.Add("TABLE_CODE1I", item.Code);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            if (reply == null)
                throw new Exception("SE HA PRODUCIDO UN ERROR AL INTENTAR CREAR EL CÓDIGO " + item.Code + " en el tipo " + item.Type);
            if(_eFunctions.CheckReplyError(reply) || _eFunctions.CheckReplyWarning(reply))
                throw new Exception(reply.message);
            if(reply.mapName != "MSM010B")
                throw new Exception("NO SE HA PODIDO CONTINUAR CON EL SIGUIENTE PASO MSM010B");

            //no hay errores ni advertencias
            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            if(replyFields.GetField("TABLE_CODE_A2I1").value != item.Code)
                throw new Exception("EL CÓDIGO MOSTRADO NO COINCIDE CON EL CÓDIGO A REGISTRAR");
            if (replyFields.GetField("TABLE_TYPE2I").value != item.Type)
                throw new Exception("EL TIPO MOSTRADO NO COINCIDE CON EL TIPO A REGISTRAR");

            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("TABLE_DESC2I1", item.Description);
            arrayFields.Add("ASSOC_REC2I1", item.AssocRec);

            request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);
            var attemp = 0;

            while (reply != null && reply.mapName == "MSM010B")
            {
                if(_eFunctions.CheckReplyError(reply))
                    throw new Exception(reply.message);
                if (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING"))
                {
                    request = new Screen.ScreenSubmitRequestDTO {screenKey = "1"};
                    reply = proxySheet.submit(opContext, request);
                }
                else
                {
                    attemp++;
                    if (attemp <= 1) continue;
                    replyFields = new ArrayScreenNameValue(reply.screenFields);
                    if (replyFields.GetField("TABLE_CODE_A2I1").value != item.Code)
                        break;
                    throw new Exception("SE HA PRODUCIDO UN ERROR AL INTENTAR CREAR EL CÓDIGO " + item.Code + " en el tipo " + item.Type);
                }
            }
        }

        public void ModifyCodeRegister(Screen.OperationContext opContext, Screen.ScreenService proxySheet, ItemCode item)
        {
            proxySheet.Url = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label) + "/ScreenService";
            _eFunctions.RevertOperation(opContext, proxySheet);
            //ejecutamos el programa
            var reply = proxySheet.executeScreen(opContext, "MSO010");
            //Validamos el ingreso
            if (reply.mapName != "MSM010A")
                throw new Exception("NO SE PUEDE INGRESAR AL PROGRAMA MSO010");
            //se adicionan los valores a los campos
            var arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("OPTION1I", "2");
            arrayFields.Add("TABLE_TYPE1I", item.Type);
            arrayFields.Add("TABLE_CODE1I", item.Code);

            var request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);

            if (reply == null)
                throw new Exception("SE HA PRODUCIDO UN ERROR AL INTENTAR CREAR EL CÓDIGO " + item.Code + " en el tipo " + item.Type);
            if (_eFunctions.CheckReplyError(reply) || _eFunctions.CheckReplyWarning(reply))
                throw new Exception(reply.message);
            if (reply.mapName != "MSM010B")
                throw new Exception("NO SE HA PODIDO CONTINUAR CON EL SIGUIENTE PASO MSM010B");

            //no hay errores ni advertencias
            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            if (replyFields.GetField("TABLE_CODE_A2I1").value != item.Code)
                throw new Exception("EL CÓDIGO MOSTRADO NO COINCIDE CON EL CÓDIGO A REGISTRAR");
            if (replyFields.GetField("TABLE_TYPE2I").value != item.Type)
                throw new Exception("EL TIPO MOSTRADO NO COINCIDE CON EL TIPO A REGISTRAR");

            arrayFields = new ArrayScreenNameValue();
            arrayFields.Add("TABLE_DESC2I1", item.Description);
            arrayFields.Add("ASSOC_REC2I1", item.AssocRec);

            request = new Screen.ScreenSubmitRequestDTO
            {
                screenFields = arrayFields.ToArray(),
                screenKey = "1"
            };
            reply = proxySheet.submit(opContext, request);
            var attemp = 0;

            while (reply != null && reply.mapName == "MSM010B")
            {
                if (_eFunctions.CheckReplyError(reply))
                    throw new Exception(reply.message);
                if (_eFunctions.CheckReplyWarning(reply) || reply.functionKeys == "XMIT-Confirm" || reply.functionKeys.StartsWith("XMIT-WARNING"))
                {
                    request = new Screen.ScreenSubmitRequestDTO { screenKey = "1" };
                    reply = proxySheet.submit(opContext, request);
                }
                else
                {
                    attemp++;
                    if (attemp <= 1) continue;
                    replyFields = new ArrayScreenNameValue(reply.screenFields);
                    if (replyFields.GetField("TABLE_CODE_A2I1").value != item.Code)
                        break;
                    throw new Exception("SE HA PRODUCIDO UN ERROR AL INTENTAR CREAR EL CÓDIGO " + item.Code + " en el tipo " + item.Type);
                }
            }
        }

        public List<ItemCode> GetItemList(EllipseFunctions ef, int searchCriteriaKey1, int typeCriteriaKey1, string searchCriteriaValue1, int searchCriteriaKey2, int typeCriteriaKey2, string searchCriteriaValue2, int statusCriteriaKey)
        {
            var sqlQuery = Queries.GetItemCodeList(ef.DbReference, ef.DbLink, searchCriteriaKey1, typeCriteriaKey1, searchCriteriaValue1, searchCriteriaKey2, typeCriteriaKey2, searchCriteriaValue2, statusCriteriaKey);
            var drItem = ef.GetQueryResult(sqlQuery);
            var list = new List<ItemCode>();

            if (drItem == null || drItem.IsClosed) return list;
            while (drItem.Read())
            {
                var order = new ItemCode
                {
                    Type = drItem["TABLE_TYPE"].ToString().Trim(),
                    TypeDescription = drItem["TYPE_DESC"].ToString().Trim(),
                    Code = drItem["TABLE_CODE"].ToString().Trim(),
                    Description = drItem["TABLE_DESC"].ToString().Trim(),
                    ActiveStatus = drItem["ACTIVE_FLAG"].ToString().Trim(),
                    AssocRec = drItem["ASSOC_REC"].ToString().Trim()
                };
                list.Add(order);
            }

            return list;
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
