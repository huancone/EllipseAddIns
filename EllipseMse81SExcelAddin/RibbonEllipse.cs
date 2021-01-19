using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Utilities;
using EllipseMse81SExcelAddin.EmployeeService;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseMse81SExcelAddin
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {


        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        Application _excelApp;
        private const string SheetName01 = "EmployeeWorkbench";
        private const string ValidationSheetName = "ValidationSheet";
        private const int TitleRow01 = 9;
        private const int ResultColumn01 = 33;
        private const string TableName01 = "WorkbenchTable";

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

        private void btnCreateEmployee_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(CreateEmployeeList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(UpdateEmployeeList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
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
                _excelApp.ActiveWorkbook.Worksheets.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "EMPLOYEE WORKBENCH MSE81S - ELLIPSE 8";
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

                for (var i = 2; i < ResultColumn01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }

                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleRequired;
                //Listas de validación
                var posReasonList = _eFunctions.GetItemCodes("TFRR");
                var valPosReasonList = posReasonList.Select(item => item.Code + " - " + item.Description).ToList();
                var persStatList = _eFunctions.GetItemCodes("EPST");
                var valPersStatList = persStatList.Select(item => item.Code + " - " + item.Description).ToList();
                var valGenderList = new List<string>{"F - Female", "M - Male", "U - Unknown"};
                var valEmpTypeList = new List<string> { "CORE", "PERS" };

                //GENERAL
                _cells.GetCell(02, TitleRow01 - 2).Value = "GENERAL";
                _cells.GetRange(02, TitleRow01 - 2, 12, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.MergeCells(02, TitleRow01 - 2, 12, TitleRow01 - 2);

                _cells.GetCell(01, TitleRow01).Value = "EMPLOYEE";
                _cells.GetCell(02, TitleRow01).Value = "EMP.TYPE";
                _cells.SetValidationList(_cells.GetCell(02, TitleRow01 +1), valEmpTypeList, ValidationSheetName, 1);

                //NAME         
                _cells.GetCell(03, TitleRow01).Value = "PREF NAME";
                _cells.GetCell(04, TitleRow01).Value = "LAST NAME";
                _cells.GetCell(05, TitleRow01).Value = "FIRST NAME";
                _cells.GetCell(06, TitleRow01).Value = "SECOND NAME";
                _cells.GetCell(06, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(07, TitleRow01).Value = "THIRD NAME";
                _cells.GetCell(07, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(08, TitleRow01).Value = "P.LAST NAME";
                _cells.GetCell(08, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(09, TitleRow01).Value = "P.FIRST NAME";
                _cells.GetCell(09, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(10, TitleRow01).Value = "P.SECOND NAME";
                _cells.GetCell(10, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(11, TitleRow01).Value = "P.THIRD NAME";
                _cells.GetCell(11, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(12, TitleRow01).Value = "GENDER";
                _cells.SetValidationList(_cells.GetCell(12, TitleRow01+1), valGenderList, ValidationSheetName, 2);
                
                
                //CONTACTS
                _cells.GetCell(13, TitleRow01 - 2).Value = "CONTACT";
                _cells.GetRange(13, TitleRow01 - 2, 19, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.MergeCells(13, TitleRow01 - 2, 19, TitleRow01 - 2);


                _cells.GetCell(13, TitleRow01).Value = "PHONE NUMBER";
                _cells.GetCell(14, TitleRow01).Value = "EXTENSION";
                _cells.GetCell(15, TitleRow01).Value = "MOBILE NUMBER";
                _cells.GetCell(16, TitleRow01).Value = "FAX NUMBER";
                _cells.GetCell(17, TitleRow01).Value = "EMAIL";
                _cells.GetCell(18, TitleRow01).Value = "MESSAGE PREF";
                _cells.GetCell(19, TitleRow01).Value = "NOTIFY EDI";
 
                _cells.GetRange(13, TitleRow01, 19 , TitleRow01).Style = StyleConstants.TitleOptional;

                //WORK DETAILS
                _cells.GetCell(20, TitleRow01 - 2).Value = "WORK DETAILS";
                _cells.GetRange(20, TitleRow01 - 2, 26, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.MergeCells(20, TitleRow01 - 2, 26, TitleRow01 - 2);

                _cells.GetCell(20, TitleRow01).Value = "PHYSICAL LOC";
                _cells.GetCell(21, TitleRow01).Value = "HIRE DATE";
                _cells.GetCell(21, TitleRow01).AddComment("YYYYMMDD");
                _cells.GetCell(22, TitleRow01).Value = "UNION CODE";
                _cells.GetCell(23, TitleRow01).Value = "WO PREFIX";
                _cells.GetCell(23, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(24, TitleRow01).Value = "RES TYPE";
                _cells.GetCell(25, TitleRow01).Value = "RES CODE";
                _cells.GetCell(26, TitleRow01).Value = "PRINTER NAME1";
                
                //POSITION
                _cells.GetCell(27, TitleRow01 - 2).Value = "POSITION";
                _cells.GetRange(27, TitleRow01 - 2, 32, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.MergeCells(27, TitleRow01 - 2, 32, TitleRow01 - 2);

                _cells.GetCell(27, TitleRow01).Value = "POSITION PRIMARY";
                _cells.GetCell(28, TitleRow01).Value = "POS REASON";
                _cells.SetValidationList(_cells.GetCell(28, TitleRow01 + 1), valPosReasonList, ValidationSheetName, 3);
                _cells.GetCell(29, TitleRow01).Value = "POS START DATE";
                _cells.GetCell(30, TitleRow01).Value = "FTE %";
                _cells.GetCell(31, TitleRow01).Value = "AUTHORITY %";
                _cells.GetCell(32, TitleRow01).Value = "PERS. STATUS";
                _cells.SetValidationList(_cells.GetCell(32, TitleRow01 + 1), valPersStatList, ValidationSheetName, 4);
                _cells.GetCell(33, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(33, TitleRow01).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja.\n" + ex);
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

        private void btnReReviewEmployees_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReReviewEmployees);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");

        }

        private void ReReviewEmployees()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);

            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstancesSpecified = true,
                maxInstances = 100,
                returnWarningsSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var proxyEmp = new EmployeeService.EmployeeService();
            
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            proxyEmp.Url = urlService + "/EmployeeService";
            
            var i = TitleRow01 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var requestReadEmp = new EmployeeServiceReadRequestDTO();
                    var attributesReadEmp = new EmployeeServiceReadRequiredAttributesDTO();
                    
                    var employee = _cells.GetEmptyIfNull(_cells.GetCell(01, i).Value);
                    
                    attributesReadEmp.returnPersonType = true;
                    attributesReadEmp.returnPreferredName = true;
                    attributesReadEmp.returnLastName = true;
                    attributesReadEmp.returnFirstName = true;
                    attributesReadEmp.returnSecondName = true;
                    attributesReadEmp.returnThirdName = true;
                    attributesReadEmp.returnPreviousLastName = true;
                    attributesReadEmp.returnPreviousFirstName = true;
                    attributesReadEmp.returnPreviousSecondName = true;
                    attributesReadEmp.returnPreviousThirdName = true;
                    attributesReadEmp.returnGender = true;
                    attributesReadEmp.returnWorkTelephoneNumber = true;
                    attributesReadEmp.returnWorkTelephoneExtension = true;
                    attributesReadEmp.returnWorkMobilePhoneNumber = true;
                    attributesReadEmp.returnHomeFacsimileNumber = true;
                    attributesReadEmp.returnEmailAddress = true;
                    attributesReadEmp.returnMessagePreference = true;
                    attributesReadEmp.returnNotifyEDIMsgRecieved = true;
                    attributesReadEmp.returnPhysicalLocation = true;
                    attributesReadEmp.returnHireDate = true;
                    attributesReadEmp.returnUnionCode = true;
                    attributesReadEmp.returnWorkOrderPrefix = true;
                    attributesReadEmp.returnResourceClass = true;
                    attributesReadEmp.returnResourceCode = true;
                    attributesReadEmp.returnPrinterName1 = true;
                    attributesReadEmp.returnPosition = true;
                    attributesReadEmp.returnPositionReason = true;
                    attributesReadEmp.returnPositionStartDate = true;
                    attributesReadEmp.returnActualFTEPercent = true;
                    attributesReadEmp.returnAuthorityPercent = true;
                    attributesReadEmp.returnPersonnelStatus = true;


                    attributesReadEmp.returnPersonTypeSpecified = true;
                    attributesReadEmp.returnPreferredNameSpecified = true;
                    attributesReadEmp.returnLastNameSpecified = true;
                    attributesReadEmp.returnFirstNameSpecified = true;
                    attributesReadEmp.returnSecondNameSpecified = true;
                    attributesReadEmp.returnThirdNameSpecified = true;
                    attributesReadEmp.returnPreviousLastNameSpecified = true;
                    attributesReadEmp.returnPreviousFirstNameSpecified = true;
                    attributesReadEmp.returnPreviousSecondNameSpecified = true;
                    attributesReadEmp.returnPreviousThirdNameSpecified = true;
                    attributesReadEmp.returnGenderSpecified = true;
                    attributesReadEmp.returnWorkTelephoneNumberSpecified = true;
                    attributesReadEmp.returnWorkTelephoneExtensionSpecified = true;
                    attributesReadEmp.returnWorkMobilePhoneNumberSpecified = true;
                    attributesReadEmp.returnHomeFacsimileNumberSpecified = true;
                    attributesReadEmp.returnEmailAddressSpecified = true;
                    attributesReadEmp.returnMessagePreferenceSpecified = true;
                    attributesReadEmp.returnNotifyEDIMsgRecievedSpecified = true;
                    attributesReadEmp.returnPhysicalLocationSpecified = true;
                    attributesReadEmp.returnHireDateSpecified = true;
                    attributesReadEmp.returnUnionCodeSpecified = true;
                    attributesReadEmp.returnWorkOrderPrefixSpecified = true;
                    attributesReadEmp.returnResourceClassSpecified = true;
                    attributesReadEmp.returnResourceCodeSpecified = true;
                    attributesReadEmp.returnPrinterName1Specified = true;
                    attributesReadEmp.returnPositionSpecified = true;
                    attributesReadEmp.returnPositionReasonSpecified = true;
                    attributesReadEmp.returnPositionStartDateSpecified = true;
                    attributesReadEmp.returnActualFTEPercentSpecified = true;
                    attributesReadEmp.returnAuthorityPercentSpecified = true;
                    attributesReadEmp.returnPersonnelStatusSpecified = true;

                    requestReadEmp.employee = employee;
                    requestReadEmp.requiredAttributes = attributesReadEmp;
                    var replyRead = proxyEmp.read(opSheet, requestReadEmp);

                    _cells.GetCell(02, i).Value = "" + replyRead.personType;                              
                    _cells.GetCell(03, i).Value = "" + replyRead.preferredName;                           
                    _cells.GetCell(04, i).Value = "" + replyRead.lastName;                                
                    _cells.GetCell(05, i).Value = "" + replyRead.firstName;                               
                    _cells.GetCell(06, i).Value = "" + replyRead.secondName;                              
                    _cells.GetCell(07, i).Value = "" + replyRead.thirdName;                               
                    _cells.GetCell(08, i).Value = "" + replyRead.previousLastName;                        
                    _cells.GetCell(09, i).Value = "" + replyRead.previousFirstName;                       
                    _cells.GetCell(10, i).Value = "" + replyRead.previousSecondName;                      
                    _cells.GetCell(11, i).Value = "" + replyRead.previousThirdName;                       
                    _cells.GetCell(12, i).Value = "" + replyRead.gender;                                  
                    _cells.GetCell(13, i).Value = "" + replyRead.workTelephoneNumber;                     
                    _cells.GetCell(14, i).Value = "" + replyRead.workTelephoneExtension;                  
                    _cells.GetCell(15, i).Value = "" + replyRead.workMobilePhoneNumber;                   
                    _cells.GetCell(16, i).Value = "" + replyRead.workFacsimileNumber;                     
                    _cells.GetCell(17, i).Value = "" + replyRead.emailAddress;                            
                    _cells.GetCell(18, i).Value = "" + replyRead.messagePreference;                       
                    _cells.GetCell(19, i).Value = "" + replyRead.notifyEDIMsgRecieved;                    
                    _cells.GetCell(20, i).Value = "" + replyRead.physicalLocation;                        
                    _cells.GetCell(21, i).Value = "" + replyRead.hireDate;                                
                    _cells.GetCell(22, i).Value = "" + replyRead.unionCode;                               
                    _cells.GetCell(23, i).Value = "" + replyRead.workOrderPrefix;                
                    _cells.GetCell(24, i).Value = "" + replyRead.resourceClass;                           
                    _cells.GetCell(25, i).Value = "" + replyRead.resourceCode;                            
                    _cells.GetCell(26, i).Value = "" + replyRead.printerName1;                            
                    _cells.GetCell(27, i).Value = "" + replyRead.position;                                
                    _cells.GetCell(28, i).Value = "" + replyRead.positionReason;                          
                    _cells.GetCell(29, i).Value = "" + replyRead.positionStartDate;                       
                    _cells.GetCell(30, i).Value = "" + replyRead.actualFTEPercent;                        
                    _cells.GetCell(31, i).Value = "" + replyRead.authorityPercent;                        
                    _cells.GetCell(32, i).Value = "" + replyRead.personnelStatus;                         


                    _cells.GetCell(ResultColumn01, i).Value = "CONSULTADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewEmployeeList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                }
            }
            _cells.SetCursorDefault();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }

        public void CreateEmployeeList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstancesSpecified = true,
                maxInstances = 100,
                returnWarningsSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var proxyEmp = new EmployeeService.EmployeeService();//ejecuta las acciones del servicio
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            proxyEmp.Url = urlService + "/EmployeeService";

            var i = TitleRow01 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var requestEmp = new EmployeeServiceCreateRequestDTO();


                    var employee = _cells.GetEmptyIfNull(_cells.GetCell(01, i).Value);
                    var employeeType = _cells.GetEmptyIfNull(_cells.GetCell(02, i).Value);//CORE, PERS
                    //NAME
                    var preferredName = _cells.GetEmptyIfNull(_cells.GetCell(03, i).Value);
                    var lastName = _cells.GetEmptyIfNull(_cells.GetCell(04, i).Value);
                    var firstName = _cells.GetEmptyIfNull(_cells.GetCell(05, i).Value);
                    var secondName = _cells.GetEmptyIfNull(_cells.GetCell(06, i).Value);
                    var thirdName = _cells.GetEmptyIfNull(_cells.GetCell(07, i).Value);
                    var prevLastName = _cells.GetEmptyIfNull(_cells.GetCell(08, i).Value);
                    var previousFirstName = _cells.GetEmptyIfNull(_cells.GetCell(09, i).Value);
                    var previousSecondName = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value);
                    var previousThirdName = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value);
                    var gender = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value);//Female, Male, Unknown
                    //CONTACTS
                    var workPhoneNumber = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value);
                    var workExtension = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value);
                    var workMobileNumber = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value);
                    var workFaxNumber = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value);
                    var workEmail = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value);
                    var messagePreference = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value);
                    var notifyEdiMsgReceived = MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(19, i).Value));
                    //WORK DETAILS
                    var physicalLocation = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value);
                    var hireDate = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value);
                    var unionCode = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value);
                    var workOrderPrefix = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value);
                    var resourceType = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value);
                    var resourceCode = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value);
                    var printerName1 = _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value);
                    //POSITION DETAILS
                    var position = _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value);
                    var positionReason = _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value);//TFRR
                    var positionStartDate = _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value);
                    var actualFtePercent = _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value);
                    var authorityPercent = _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value);
                    var personnelStatus = _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value);//EPST
                    

                    requestEmp.employee = employee;
                    requestEmp.personType = MyUtilities.GetCodeKey(employeeType);

                    requestEmp.preferredName = preferredName;
                    requestEmp.lastName = lastName;
                    requestEmp.firstName = firstName;
                    requestEmp.secondName = secondName;
                    requestEmp.thirdName = thirdName;
                    requestEmp.previousLastName = prevLastName;
                    requestEmp.previousFirstName = previousFirstName;
                    requestEmp.previousSecondName = previousSecondName;
                    requestEmp.previousThirdName = previousThirdName;
                    requestEmp.gender = MyUtilities.GetCodeKey(gender);

                    requestEmp.workTelephoneNumber = workPhoneNumber;
                    requestEmp.workTelephoneExtension = workExtension;
                    requestEmp.workMobilePhoneNumber = workMobileNumber;
                    requestEmp.workFacsimileNumber = workFaxNumber;
                    requestEmp.emailAddress = workEmail;
                    requestEmp.messagePreference = messagePreference;
                    requestEmp.notifyEDIMsgRecievedSpecified = true;
                    requestEmp.notifyEDIMsgRecieved = notifyEdiMsgReceived;

                    requestEmp.physicalLocation = physicalLocation;
                    requestEmp.hireDate = hireDate;
                    requestEmp.unionCode = unionCode;
                    requestEmp.workOrderPrefix = workOrderPrefix;
                    requestEmp.resourceClass = resourceType;
                    requestEmp.resourceCode = resourceCode;
                    requestEmp.printerName1 = printerName1;

                    requestEmp.position = position;
                    requestEmp.positionReason = MyUtilities.GetCodeKey(positionReason);
                    requestEmp.positionStartDate = positionStartDate;
                    requestEmp.actualFTEPercent = !string.IsNullOrWhiteSpace(actualFtePercent) ? Convert.ToDecimal(actualFtePercent) : default(decimal);
                    requestEmp.actualFTEPercentSpecified = true;
                    requestEmp.authorityPercent = !string.IsNullOrWhiteSpace(authorityPercent) ? Convert.ToDecimal(authorityPercent) : default(decimal);
                    requestEmp.authorityPercentSpecified = true;
                    requestEmp.personnelStatus = MyUtilities.GetCodeKey(personnelStatus);

                    

                    proxyEmp.create(opSheet, requestEmp);


                    _cells.GetCell(ResultColumn01, i).Value = "CREADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateEmployeeList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                }
            }
            _cells.SetCursorDefault();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }

        public void UpdateEmployeeList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            _cells.SetCursorWait();
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstancesSpecified = true,
                maxInstances = 100,
                returnWarningsSpecified = true,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var proxyEmp = new EmployeeService.EmployeeService();//ejecuta las acciones del servicio
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            proxyEmp.Url = urlService + "/EmployeeService";
            
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var requestEmp = new EmployeeServiceModifyRequestDTO();

                    var employee = _cells.GetEmptyIfNull(_cells.GetCell(01, i).Value);
                    var employeeType = MyUtilities.IsTrue(_cells.GetCell(02, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(02, i).Value) : null;//CORE, PERS
                    //NAME
                    var preferredName = MyUtilities.IsTrue(_cells.GetCell(03, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(03, i).Value) : null;
                    var lastName = MyUtilities.IsTrue(_cells.GetCell(04, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(04, i).Value) : null;
                    var firstName = MyUtilities.IsTrue(_cells.GetCell(05, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(05, i).Value) : null;
                    var secondName = MyUtilities.IsTrue(_cells.GetCell(06, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(06, i).Value) : null;
                    var thirdName = MyUtilities.IsTrue(_cells.GetCell(07, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(07, i).Value) : null;
                    var prevLastName = MyUtilities.IsTrue(_cells.GetCell(08, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(08, i).Value) : null;
                    var previousFirstName = MyUtilities.IsTrue(_cells.GetCell(09, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(09, i).Value) : null;
                    var previousSecondName = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null;
                    var previousThirdName = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    var gender = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;//Female, Male, Unknown
                    //CONTACTS
                    var workPhoneNumber = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    var workExtension = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null;
                    var workMobileNumber = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    var workFaxNumber = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    var workEmail = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null;
                    var messagePreference = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;
                    var notifyEdiMsgReceived = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? (_cells.GetCell(19, i).Value) : null;
                    var notifyEdiMsgReceivedSpecified = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value);
                    //WORK DETAILS
                    var physicalLocation = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null;
                    var hireDate = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null;
                    var unionCode = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    var workOrderPrefix = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null;
                    var resourceType = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    var resourceCode = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null;
                    var printerName1 = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null;
                    //POSITION DETAILS
                    var position = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null;
                    var positionReason = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null;//TFRR
                    var positionStartDate = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;
                    var actualFtePercent = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null;
                    var actualFtePercentSpecified = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value);
                    var authorityPercent = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null;
                    var authorityPercentSpecified = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value);
                    var personnelStatus = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value) ?_cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null;//EPST


                    requestEmp.employee = employee;
                    requestEmp.personType = employeeType;

                    requestEmp.preferredName = preferredName;
                    requestEmp.lastName = lastName;
                    requestEmp.firstName = firstName;
                    requestEmp.secondName = secondName;
                    requestEmp.thirdName = thirdName;
                    requestEmp.previousLastName = prevLastName;
                    requestEmp.previousFirstName = previousFirstName;
                    requestEmp.previousSecondName = previousSecondName;
                    requestEmp.previousThirdName = previousThirdName;
                    requestEmp.gender = MyUtilities.GetCodeKey(gender);

                    requestEmp.workTelephoneNumber = workPhoneNumber;
                    requestEmp.workTelephoneExtension = workExtension;
                    requestEmp.workMobilePhoneNumber = workMobileNumber;
                    requestEmp.workFacsimileNumber = workFaxNumber;
                    requestEmp.emailAddress = workEmail;
                    requestEmp.messagePreference = messagePreference;
                    requestEmp.notifyEDIMsgRecievedSpecified = notifyEdiMsgReceivedSpecified;
                    requestEmp.notifyEDIMsgRecieved = MyUtilities.IsTrue(notifyEdiMsgReceived);

                    requestEmp.physicalLocation = physicalLocation;
                    requestEmp.hireDate = hireDate;
                    requestEmp.unionCode = unionCode;
                    requestEmp.workOrderPrefix = workOrderPrefix;
                    requestEmp.resourceClass = resourceType;
                    requestEmp.resourceCode = resourceCode;
                    requestEmp.printerName1 = printerName1;

                    requestEmp.position = position;
                    requestEmp.positionReason = MyUtilities.GetCodeKey(positionReason);
                    requestEmp.positionStartDate = positionStartDate;
                    requestEmp.actualFTEPercentSpecified = actualFtePercentSpecified;
                    requestEmp.actualFTEPercent = actualFtePercentSpecified ? Convert.ToDecimal(actualFtePercent) : 0;
                    requestEmp.authorityPercentSpecified = authorityPercentSpecified;
                    requestEmp.authorityPercent = authorityPercentSpecified ? Convert.ToDecimal(authorityPercent) : 0;
                    requestEmp.personnelStatus = MyUtilities.GetCodeKey(personnelStatus);

                    proxyEmp.modify(opSheet, requestEmp);


                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateEmployeeList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                }
            }
            _cells.SetCursorDefault();
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        

    }
}
