using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Web.Services.Ellipse;
using EllipseEquipmentClassLibrary;
using EquipmentService = EllipseEquipmentClassLibrary.EquipmentService;
using EquipTraceService = EllipseEquipmentClassLibrary.EquipTraceService;
using System.Threading;

namespace EllipseEquipmentExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;
        private const string SheetName01 = "EquipmentFull";
        private const string SheetName02 = "TracingActions";
        private const string ValidationSheetName = "ValidationSheetEquipment";

        private const int TitleRow01 = 9;
        private const int TitleRow02 = 8;
        private const int ResultColumn01 = 73;
        private const int ResultColumn02 = 9;
        private const string TableName01 = "EquipmentTable";
        private const string TableName02 = "FitmentDefitmentTable";

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
        private void btnFormatFull_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }
        private void btnCreateEquipment_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(CreateEquipment);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnUpdateEquipmentData_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(UpdateEquipment);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnUpdateEquipmentStatus_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(UpdateEquipmentStatus);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnDeleteEquipment_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeleteEquipment);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnDefitment_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ExecuteTraceAction);

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
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.CreateNewWorksheet(ValidationSheetName);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "EQUIPMENT - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, TitleRow01-2, ResultColumn01 - 1, TitleRow01-2).Style = StyleConstants.Select;
                _cells.GetRange(1, TitleRow01, ResultColumn01-1, TitleRow01).Style = StyleConstants.TitleOptional;
                for (var i = 2; i < ResultColumn01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1)
                        .AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }


                //GENERAL
                _cells.GetCell(1, TitleRow01 - 2).Value = "GENERAL";
                _cells.MergeCells(1, TitleRow01 - 2, 23, TitleRow01 - 2);
                _cells.GetCell(1, TitleRow01).Value = "EQUIPMENT REFERENCE";
                _cells.GetCell(2, TitleRow01).Value = "STATUS";//ES
                _cells.GetCell(3, TitleRow01).Value = "DESCRIPTION 1";
                _cells.GetCell(3, TitleRow01).AddComment("40 Caracteres");
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPTION 2";
                _cells.GetCell(4, TitleRow01).AddComment("40 Caracteres");
                _cells.GetCell(5, TitleRow01).Value = "EQ. CLASS";//EC
                _cells.GetRange(1, TitleRow01, 5, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(6, TitleRow01).Value = "EQ. TYPE";//ET
                _cells.GetCell(7, TitleRow01).Value = "PLANT NO";
                _cells.GetCell(8, TitleRow01).Value = "EGI";
                _cells.GetCell(9, TitleRow01).Value = "EQ. LOCATION";//EQ
                _cells.GetCell(10, TitleRow01).Value = "PRODUCTIVE UNIT";
                _cells.GetCell(11, TitleRow01).Value = "CUSTODIAN";
                _cells.GetCell(12, TitleRow01).Value = "CUST. POS.";
                _cells.GetCell(13, TitleRow01).Value = "OPERATOR";
                _cells.GetCell(14, TitleRow01).Value = "OPRT. POS.";
                _cells.GetCell(15, TitleRow01).Value = "INPUT BY";
                _cells.GetCell(16, TitleRow01).Value = "CUSTOMER NUMBER";
                _cells.GetCell(17, TitleRow01).Value = "SHUTDOWN EQUIP";
                _cells.GetCell(18, TitleRow01).Value = "WARRANTY TYPE";//SS
                _cells.GetCell(19, TitleRow01).Value = "WAR. STAT. VALUE";
                _cells.GetCell(20, TitleRow01).Value = "WAR. DATE";
                _cells.GetCell(21, TitleRow01).Value = "NAME CODE";
                _cells.GetCell(22, TitleRow01).Value = "DISTRICT CODE";
                _cells.GetCell(23, TitleRow01).Value = "ACTIVE FLAG";
                _cells.GetCell(23, TitleRow01).AddComment("Y/N");
                _cells.GetRange(22, TitleRow01, 23, TitleRow01).Style = StyleConstants.TitleRequired;
                //COSTING
                _cells.GetCell(24, TitleRow01 - 2).Value = "COSTING";
                _cells.MergeCells(24, TitleRow01 - 2, 33, TitleRow01 - 2);
                _cells.GetCell(24, TitleRow01).Value = "ACCOUNT CODE";
                _cells.GetCell(25, TitleRow01).Value = "EXP. ELEMENT";
                _cells.GetCell(26, TitleRow01).Value = "COST FLAG";
                _cells.GetCell(26, TitleRow01).AddComment("A - Allowed, W - Not Allowed (Warning), E - Not Allowed (Error)");
                _cells.GetCell(27, TitleRow01).Value = "TAX CODE";
                _cells.GetRange(24, TitleRow01, 27, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(28, TitleRow01).Value = "CONSMP. TAX CODE";
                _cells.GetCell(29, TitleRow01).Value = "PURCH. ORDER";
                _cells.GetCell(30, TitleRow01).Value = "PURCH. DATE";
                _cells.GetCell(31, TitleRow01).Value = "PURCH. PRICE";
                _cells.GetCell(32, TitleRow01).Value = "REPL. VALUE";
                _cells.GetCell(33, TitleRow01).Value = "VALUATION DATE";
                //TRACING
                _cells.GetCell(34, TitleRow01 - 2).Value = "TRACING";
                _cells.MergeCells(34, TitleRow01 - 2, 41, TitleRow01 - 2);
                _cells.GetCell(34, TitleRow01).Value = "COMP. CODE";//CO
                _cells.GetCell(35, TitleRow01).Value = "MNEMONIC";//AA
                _cells.GetCell(36, TitleRow01).Value = "STOCK CODE";
                _cells.GetCell(37, TitleRow01).Value = "SERIAL NUMBER";
                _cells.GetCell(38, TitleRow01).Value = "PART NUMBER";
                _cells.GetCell(39, TitleRow01).Value = "DOC. NUMBER";
                _cells.GetCell(40, TitleRow01).Value = "ORIG. DOCUMENT";
                _cells.GetCell(41, TitleRow01).Value = "TRACEABLE FLAG";
                //CONDITION
                _cells.GetCell(42, TitleRow01 - 2).Value = "CONDITION";
                _cells.MergeCells(42, TitleRow01 - 2, 48, TitleRow01 - 2);
                _cells.GetCell(42, TitleRow01).Value = "EQ. CRITICALLY";//EQCR
                _cells.GetCell(43, TitleRow01).Value = "PRIMARY FUNCTION";
                _cells.GetCell(44, TitleRow01).Value = "OPERATING STANDARD";
                _cells.GetCell(45, TitleRow01).Value = "COND. STANDARD";//EQCN
                _cells.GetCell(46, TitleRow01).Value = "COND. RATING";
                _cells.GetCell(47, TitleRow01).Value = "LATEST COND. DATE";
                _cells.GetCell(48, TitleRow01).Value = "MSSS APPLIES FLAG";
                //LINK ONE
                _cells.GetCell(49, TitleRow01 - 2).Value = "LINK ONE";
                _cells.MergeCells(49, TitleRow01 - 2, 52, TitleRow01 - 2);
                _cells.GetCell(49, TitleRow01).Value = "PUBLISHER";
                _cells.GetCell(50, TitleRow01).Value = "BOOK";
                _cells.GetCell(51, TitleRow01).Value = "PAGE REF.";
                _cells.GetCell(52, TitleRow01).Value = "ITEM ID.";

                _cells.GetRange(49, TitleRow01, 52, TitleRow01).Style = StyleConstants.TitleInformation;
                //CLASSIFICATION CODES
                _cells.GetCell(53, TitleRow01 - 2).Value = "CLASSIFICATION CODES";
                _cells.MergeCells(53, TitleRow01 - 2, 72, TitleRow01 - 2);
                _cells.GetCell(53, TitleRow01).Value = "E00. CRITICALITY";//E0
                _cells.GetCell(54, TitleRow01).Value = "E01. MT.POLICY";//E1
                _cells.GetCell(55, TitleRow01).Value = "E02. EQ.CLASSIF";//E2
                _cells.GetCell(56, TitleRow01).Value = "E03. OWNERSHIP";//E3
                _cells.GetCell(57, TitleRow01).Value = "E04. CONDITION";//E4
                _cells.GetCell(58, TitleRow01).Value = "E05. EQC5";//E5
                _cells.GetCell(59, TitleRow01).Value = "E06. EQC6";//E6
                _cells.GetCell(60, TitleRow01).Value = "E07. EQC7";//E7
                _cells.GetCell(61, TitleRow01).Value = "E08. EQC8";//E8
                _cells.GetCell(62, TitleRow01).Value = "E09. EQC9";//E9
                _cells.GetCell(63, TitleRow01).Value = "E10. MODEL0";//E10
                _cells.GetCell(64, TitleRow01).Value = "E11. MODEL1";//E11
                _cells.GetCell(65, TitleRow01).Value = "E12. MODEL2";//E12
                _cells.GetCell(66, TitleRow01).Value = "E13. MODEL3";//E14
                _cells.GetCell(67, TitleRow01).Value = "E14. MODEL4";//E14
                _cells.GetCell(68, TitleRow01).Value = "E15. MODEL5";//E15
                _cells.GetCell(69, TitleRow01).Value = "E16. MODEL6";//E16
                _cells.GetCell(70, TitleRow01).Value = "E17. MODEL7";//E17
                _cells.GetCell(71, TitleRow01).Value = "E18. MODEL8";//E18
                _cells.GetCell(72, TitleRow01).Value = "E19. MODEL9";//E19
                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);

                //validationRows
                
                //asigno la validación de celda
                var validList = DistrictConstants.GetDistrictList();
                _cells.SetValidationList(_cells.GetCell(5, TitleRow01 + 1), validList, ValidationSheetName, 1);

                var equipmentStatusList = EquipmentActions.GetEquipmentStatusCodeList(_eFunctions);
                var statusList = equipmentStatusList.Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(2, TitleRow01 + 1), statusList, ValidationSheetName, 2, false);

                validList = _eFunctions.GetItemCodes("EC").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(5, TitleRow01 + 1), validList, ValidationSheetName, 3, false);
                validList = _eFunctions.GetItemCodes("ET").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(6, TitleRow01 + 1), validList, ValidationSheetName, 4, false);
                validList = _eFunctions.GetItemCodes("EL").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(9, TitleRow01 + 1), validList, ValidationSheetName, 5, false);
                validList = _eFunctions.GetItemCodes("SS").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(18, TitleRow01 + 1), validList, ValidationSheetName, 6, false);
                validList = _eFunctions.GetItemCodes("CO").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(34, TitleRow01 + 1), validList, ValidationSheetName, 7, false);
                validList = _eFunctions.GetItemCodes("AA").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(35, TitleRow01 + 1), validList, ValidationSheetName, 8, false);
                validList = _eFunctions.GetItemCodes("EQCR").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(42, TitleRow01 + 1), validList, ValidationSheetName, 9, false);
                validList = _eFunctions.GetItemCodes("EQCN").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(45, TitleRow01 + 1), validList, ValidationSheetName, 10, false);
                //validación de celda - classification codes
                validList = _eFunctions.GetItemCodes("E0").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(53, TitleRow01 + 1), validList, ValidationSheetName, 11, false);
                validList = _eFunctions.GetItemCodes("E1").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(54, TitleRow01 + 1), validList, ValidationSheetName, 12, false);
                validList = _eFunctions.GetItemCodes("E2").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(55, TitleRow01 + 1), validList, ValidationSheetName, 13, false);
                validList = _eFunctions.GetItemCodes("E3").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(56, TitleRow01 + 1), validList, ValidationSheetName, 14, false);
                validList = _eFunctions.GetItemCodes("E4").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(57, TitleRow01 + 1), validList, ValidationSheetName, 15, false);
                validList = _eFunctions.GetItemCodes("E5").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(58, TitleRow01 + 1), validList, ValidationSheetName, 16, false);
                validList = _eFunctions.GetItemCodes("E6").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(59, TitleRow01 + 1), validList, ValidationSheetName, 17, false);
                validList = _eFunctions.GetItemCodes("E7").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(60, TitleRow01 + 1), validList, ValidationSheetName, 18, false);
                validList = _eFunctions.GetItemCodes("E8").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(61, TitleRow01 + 1), validList, ValidationSheetName, 19, false);
                validList = _eFunctions.GetItemCodes("E9").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(62, TitleRow01 + 1), validList, ValidationSheetName, 20, false);
                validList = _eFunctions.GetItemCodes("E10").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(63, TitleRow01 + 1), validList, ValidationSheetName, 21, false);
                validList = _eFunctions.GetItemCodes("E11").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(64, TitleRow01 + 1), validList, ValidationSheetName, 22, false);
                validList = _eFunctions.GetItemCodes("E12").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(65, TitleRow01 + 1), validList, ValidationSheetName, 23, false);
                validList = _eFunctions.GetItemCodes("E13").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(66, TitleRow01 + 1), validList, ValidationSheetName, 24, false);
                validList = _eFunctions.GetItemCodes("E14").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(67, TitleRow01 + 1), validList, ValidationSheetName, 25, false);
                validList = _eFunctions.GetItemCodes("E15").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(68, TitleRow01 + 1), validList, ValidationSheetName, 26, false);
                validList = _eFunctions.GetItemCodes("E16").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(69, TitleRow01 + 1), validList, ValidationSheetName, 27, false);
                validList = _eFunctions.GetItemCodes("E17").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(70, TitleRow01 + 1), validList, ValidationSheetName, 28, false);
                validList = _eFunctions.GetItemCodes("E18").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(71, TitleRow01 + 1), validList, ValidationSheetName, 29, false);
                validList = _eFunctions.GetItemCodes("E19").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(72, TitleRow01 + 1), validList, ValidationSheetName, 30, false);
                //
                validList = new List<string>() { "A - Allowed", "W - Not Allowed (Warning)", "E - Not Allowed (Error)" };
                _cells.SetValidationList(_cells.GetCell(26, TitleRow01 + 1), validList, ValidationSheetName, 31, false);
                validList = _eFunctions.GetItemCodes("TC").Select(sc => sc.code + " - " + sc.description).ToList();
                _cells.SetValidationList(_cells.GetCell(27, TitleRow01 + 1), validList, ValidationSheetName, 32, false);


                744

                //HOJA 2 - TRACING ACTIONS
                _excelApp.ActiveWorkbook.Sheets[2].Select(Type.Missing);

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "EQUIPMENT TRACING - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell(1, TitleRow02).Value = "EQUIPMENT REFERENCE";
                _cells.GetCell(2, TitleRow02).Value = "EQ. DESCRIPTION 1";
                _cells.GetCell(3, TitleRow02).Value = "EQ. DESCRIPTION 2";
                _cells.GetCell(4, TitleRow02).Value = "EGI";

                _cells.GetRange(1, TitleRow02 - 1, ResultColumn02 - 1, TitleRow02 - 1).Style = StyleConstants.Select;
                //INSTALL POSITION
                _cells.GetCell(1, TitleRow02 - 1).Value = "INSTALL POSITION";
                _cells.MergeCells(1, TitleRow02 - 1, 3, TitleRow02 - 1);

                _cells.GetCell(1, TitleRow02).Value = "INSTALL EQ. REF.";
                _cells.GetCell(2, TitleRow02).Value = "COMP CODE";
                _cells.GetCell(3, TitleRow02).Value = "MOD CODE";
                _cells.GetRange(1, TitleRow02, 3, TitleRow02).Style = StyleConstants.TitleRequired;
                //ACTION
                _cells.GetCell(4, TitleRow02 - 1).Value = "ACTION";
                _cells.GetCell(4, TitleRow02).Value = "F/D";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleAction;
                //FITTED EQUIPMENT
                _cells.GetCell(5, TitleRow02 - 1).Value = "INSTALL POSITION";
                _cells.MergeCells(5, TitleRow02 - 1, 8, TitleRow02 - 1);

                _cells.GetCell(5, TitleRow02).Value = "FIT EQ. REF";
                _cells.GetCell(6, TitleRow02).Value = "DATE";
                _cells.GetCell(7, TitleRow02).Value = "REF. TYPE";
                _cells.GetCell(8, TitleRow02).Value = "REF. NUMBER";
                _cells.GetRange(4, TitleRow02, 8, TitleRow02).Style = StyleConstants.TitleRequired;


                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = _cells.GetStyle(StyleConstants.TitleResult);

                //validationRows

                //asigno la validación de celda
                var listTypeAction = new List<string> { "F - Fitment", "D - Defitment" };
                _cells.SetValidationList(_cells.GetCell(4, TitleRow02 + 1), listTypeAction, ValidationSheetName, 1, false);

                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
                //búsquedas especiales de tabla
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        
        public void CreateEquipment()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;
            var opSheet = new EquipmentService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = _eFunctions.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    var equipment = new Equipment
                    {
                        EquipmentNo = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value),
                        EquipmentStatus = Utils.IsTrue(_cells.GetCell(2, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value)) : null,
                        EquipmentNoDescription1 = Utils.IsTrue(_cells.GetCell(3, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value) : null,
                        EquipmentNoDescription2 = Utils.IsTrue(_cells.GetCell(4, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(4, i).Value) : null,
                        EquipmentClass = Utils.IsTrue(_cells.GetCell(5, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(5, i).Value)) : null,
                        EquipmentType = Utils.IsTrue(_cells.GetCell(6, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(6, i).Value)) : null,
                        PlantNo = Utils.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(7, i).Value) : null,
                        EquipmentGrpId = Utils.IsTrue(_cells.GetCell(8, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(8, i).Value) : null,
                        Location = Utils.IsTrue(_cells.GetCell(9, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(9, i).Value)) : null,
                        ParentEquipmentRef = Utils.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(10, i).Value) : null,
                        Custodian = Utils.IsTrue(_cells.GetCell(11, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(11, i).Value)) : null,
                        CustodianPosition = Utils.IsTrue(_cells.GetCell(12, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(12, i).Value)) : null,
                        OperatorId = Utils.IsTrue(_cells.GetCell(13, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(13, i).Value)) : null,
                        OperatorPosition = Utils.IsTrue(_cells.GetCell(14, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(14, i).Value)) : null,
                        InputBy = Utils.IsTrue(_cells.GetCell(15, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(15, i).Value)) : null,
                        CustomerNumber = Utils.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(16, i).Value) : null,
                        ShutdownEquipment = Utils.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(17, i).Value) : null,
                        WarrStatType = Utils.IsTrue(_cells.GetCell(18, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(18, i).Value)) : null,
                        WarrStatVal = Utils.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(19, i).Value) : null,
                        WarrantyDate = Utils.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(20, i).Value) : null,
                        ItemNameCode = Utils.IsTrue(_cells.GetCell(21, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(21, i).Value)) : null,
                        DistrictCode = Utils.IsTrue(_cells.GetCell(22, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(22, i).Value)) : null,
                        ActiveFlag = Utils.IsTrue(_cells.GetCell(23, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(23, i).Value)) : null,
                        AccountCode = Utils.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(24, i).Value) : null,
                        ExpElement = Utils.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(25, i).Value) : null,
                        CostingFlag = Utils.IsTrue(_cells.GetCell(26, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(26, i).Value)) : null,
                        TaxCode = Utils.IsTrue(_cells.GetCell(27, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(27, i).Value)) : null,
                        CtaxCode = Utils.IsTrue(_cells.GetCell(28, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(28, i).Value)) : null,
                        PoNo = Utils.IsTrue(_cells.GetCell(29, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(29, i).Value)) : null,
                        PurchaseDate = Utils.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(30, i).Value) : null,
                        PurchasePrice = Utils.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(31, i).Value) : null,
                        ReplaceValue = Utils.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(32, i).Value) : null,
                        ValuationDate = Utils.IsTrue(_cells.GetCell(33, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(33, i).Value)) : null,
                        CompCode = Utils.IsTrue(_cells.GetCell(34, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(34, i).Value)) : null,
                        Mnemonic = Utils.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(35, i).Value) : null,
                        StockCode = Utils.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(36, i).Value) : null,
                        SerialNumber = Utils.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(37, i).Value) : null,
                        PartNo = Utils.IsTrue(_cells.GetCell(38, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(38, i).Value) : null,
                        DrawingNo = Utils.IsTrue(_cells.GetCell(39, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(39, i).Value) : null,
                        OriginalDoc = Utils.IsTrue(_cells.GetCell(40, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(40, i).Value) : null,
                        TraceableFlg = Utils.IsTrue(_cells.GetCell(41, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(41, i).Value)) : null,
                        EquipmentCriticality = Utils.IsTrue(_cells.GetCell(42, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(42, i).Value)) : null,
                        PrimaryFunction = Utils.IsTrue(_cells.GetCell(43, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(43, i).Value)) : null,
                        OperatingStandard = Utils.IsTrue(_cells.GetCell(44, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(44, i).Value)) : null,
                        ConditionStandard = Utils.IsTrue(_cells.GetCell(45, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(45, i).Value)) : null,
                        ConditionRating = Utils.IsTrue(_cells.GetCell(46, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(46, i).Value)) : null,
                        LatestConditionDate = Utils.IsTrue(_cells.GetCell(47, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(47, i).Value)) : null,
                        MsssFlag = Utils.IsTrue(_cells.GetCell(48, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(48, i).Value)) : null,
                        LinkOne = new Equipment.LinkOneBook
                        {
                            Publisher = Utils.IsTrue(_cells.GetCell(49, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(49, i).Value) : null,
                            Book = Utils.IsTrue(_cells.GetCell(50, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(50, i).Value) : null,
                            PageReference = Utils.IsTrue(_cells.GetCell(51, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(51, i).Value) : null,
                            ItemId = Utils.IsTrue(_cells.GetCell(51, validationRow).Value) ? _cells.GetNullIfTrimmedEmpty(_cells.GetCell(52, i).Value) : null
                        },
                        ClassCodes = new Equipment.ClassificationCodes()
                        {
                            EquipmentClassif0 = Utils.IsTrue(_cells.GetCell(53, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(53, i).Value)) : null,
                            EquipmentClassif1 = Utils.IsTrue(_cells.GetCell(54, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(54, i).Value)) : null,
                            EquipmentClassif2 = Utils.IsTrue(_cells.GetCell(55, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(55, i).Value)) : null,
                            EquipmentClassif3 = Utils.IsTrue(_cells.GetCell(56, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(56, i).Value)) : null,
                            EquipmentClassif4 = Utils.IsTrue(_cells.GetCell(57, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(57, i).Value)) : null,
                            EquipmentClassif5 = Utils.IsTrue(_cells.GetCell(58, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(58, i).Value)) : null,
                            EquipmentClassif6 = Utils.IsTrue(_cells.GetCell(59, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(59, i).Value)) : null,
                            EquipmentClassif7 = Utils.IsTrue(_cells.GetCell(60, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(60, i).Value)) : null,
                            EquipmentClassif8 = Utils.IsTrue(_cells.GetCell(61, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(61, i).Value)) : null,
                            EquipmentClassif9 = Utils.IsTrue(_cells.GetCell(62, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(62, i).Value)) : null,
                            EquipmentClassif10 = Utils.IsTrue(_cells.GetCell(63, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(63, i).Value)) : null,
                            EquipmentClassif11 = Utils.IsTrue(_cells.GetCell(64, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(64, i).Value)) : null,
                            EquipmentClassif12 = Utils.IsTrue(_cells.GetCell(65, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(65, i).Value)) : null,
                            EquipmentClassif13 = Utils.IsTrue(_cells.GetCell(66, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(66, i).Value)) : null,
                            EquipmentClassif14 = Utils.IsTrue(_cells.GetCell(67, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(67, i).Value)) : null,
                            EquipmentClassif15 = Utils.IsTrue(_cells.GetCell(68, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(68, i).Value)) : null,
                            EquipmentClassif16 = Utils.IsTrue(_cells.GetCell(69, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(69, i).Value)) : null,
                            EquipmentClassif17 = Utils.IsTrue(_cells.GetCell(70, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(70, i).Value)) : null,
                            EquipmentClassif18 = Utils.IsTrue(_cells.GetCell(71, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(71, i).Value)) : null,
                            EquipmentClassif19 = Utils.IsTrue(_cells.GetCell(72, validationRow).Value) ? Utils.GetCodeKey(_cells.GetNullIfTrimmedEmpty(_cells.GetCell(72, i).Value)) : null
                        }
                    };

                    EquipmentActions.CreateEquipment(opSheet, urlService, equipment);

                    _cells.GetCell(ResultColumn01, i).Value = "CREADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn01, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:UpdateEquipment()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void UpdateEquipment()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            var opSheet = new EquipmentService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = _eFunctions.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

                    var equipmentRef = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);

                    List<string> equipmentList = EquipmentActions.GetEquipmentList(_eFunctions, opSheet.district, equipmentRef);
                    var equipment = new Equipment
                    {
                        EquipmentNo = equipmentList.Any() ? equipmentList.First() : equipmentRef,
                        EquipmentStatus = Utils.IsTrue(_cells.GetCell(2, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(2, i).Value)) : null,
                        EquipmentNoDescription1 = Utils.IsTrue(_cells.GetCell(3, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value) : null,
                        EquipmentNoDescription2 = Utils.IsTrue(_cells.GetCell(4, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value) : null,
                        EquipmentClass = Utils.IsTrue(_cells.GetCell(5, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(5, i).Value)) : null,
                        EquipmentType = Utils.IsTrue(_cells.GetCell(6, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(6, i).Value)) : null,
                        PlantNo = Utils.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null,
                        EquipmentGrpId = Utils.IsTrue(_cells.GetCell(8, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value) : null,
                        Location = Utils.IsTrue(_cells.GetCell(9, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(9, i).Value)) : null,
                        ParentEquipmentRef = Utils.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null,
                        Custodian = Utils.IsTrue(_cells.GetCell(11, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value)) : null,
                        CustodianPosition = Utils.IsTrue(_cells.GetCell(12, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value)) : null,
                        OperatorId = Utils.IsTrue(_cells.GetCell(13, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(13, i).Value)) : null,
                        OperatorPosition = Utils.IsTrue(_cells.GetCell(14, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)) : null,
                        InputBy = Utils.IsTrue(_cells.GetCell(15, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(15, i).Value)) : null,
                        CustomerNumber = Utils.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null,
                        ShutdownEquipment = Utils.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null,
                        WarrStatType = Utils.IsTrue(_cells.GetCell(18, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(18, i).Value)) : null,
                        WarrStatVal = Utils.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null,
                        WarrantyDate = Utils.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null,
                        ItemNameCode = Utils.IsTrue(_cells.GetCell(21, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(21, i).Value)) : null,
                        DistrictCode = Utils.IsTrue(_cells.GetCell(22, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(22, i).Value)) : null,
                        ActiveFlag = Utils.IsTrue(_cells.GetCell(23, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(23, i).Value)) : null,
                        AccountCode = Utils.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null,
                        ExpElement = Utils.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null,
                        CostingFlag = Utils.IsTrue(_cells.GetCell(26, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(26, i).Value)) : null,
                        TaxCode = Utils.IsTrue(_cells.GetCell(27, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(27, i).Value)) : null,
                        CtaxCode = Utils.IsTrue(_cells.GetCell(28, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(28, i).Value)) : null,
                        PoNo = Utils.IsTrue(_cells.GetCell(29, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(29, i).Value)) : null,
                        PurchaseDate = Utils.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null,
                        PurchasePrice = Utils.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null,
                        ReplaceValue = Utils.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null,
                        ValuationDate = Utils.IsTrue(_cells.GetCell(33, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(33, i).Value)) : null,
                        CompCode = Utils.IsTrue(_cells.GetCell(34, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(34, i).Value)) : null,
                        Mnemonic = Utils.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null,
                        StockCode = Utils.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null,
                        SerialNumber = Utils.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null,
                        PartNo = Utils.IsTrue(_cells.GetCell(38, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value) : null,
                        DrawingNo = Utils.IsTrue(_cells.GetCell(39, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value) : null,
                        OriginalDoc = Utils.IsTrue(_cells.GetCell(40, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(40, i).Value) : null,
                        TraceableFlg = Utils.IsTrue(_cells.GetCell(41, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(41, i).Value)) : null,
                        EquipmentCriticality = Utils.IsTrue(_cells.GetCell(42, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(42, i).Value)) : null,
                        PrimaryFunction = Utils.IsTrue(_cells.GetCell(43, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(43, i).Value)) : null,
                        OperatingStandard = Utils.IsTrue(_cells.GetCell(44, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(44, i).Value)) : null,
                        ConditionStandard = Utils.IsTrue(_cells.GetCell(45, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(45, i).Value)) : null,
                        ConditionRating = Utils.IsTrue(_cells.GetCell(46, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(46, i).Value)) : null,
                        LatestConditionDate = Utils.IsTrue(_cells.GetCell(47, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(47, i).Value)) : null,
                        MsssFlag = Utils.IsTrue(_cells.GetCell(48, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(48, i).Value)) : null,
                        LinkOne = new Equipment.LinkOneBook
                        {
                            Publisher = Utils.IsTrue(_cells.GetCell(49, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(49, i).Value) : null,
                            Book = Utils.IsTrue(_cells.GetCell(50, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(50, i).Value) : null,
                            PageReference = Utils.IsTrue(_cells.GetCell(51, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(51, i).Value) : null,
                            ItemId = Utils.IsTrue(_cells.GetCell(51, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(52, i).Value) : null
                        },
                        ClassCodes = new Equipment.ClassificationCodes()
                        {
                            EquipmentClassif0 = Utils.IsTrue(_cells.GetCell(53, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(53, i).Value)) : null,
                            EquipmentClassif1 = Utils.IsTrue(_cells.GetCell(54, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(54, i).Value)) : null,
                            EquipmentClassif2 = Utils.IsTrue(_cells.GetCell(55, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(55, i).Value)) : null,
                            EquipmentClassif3 = Utils.IsTrue(_cells.GetCell(56, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(56, i).Value)) : null,
                            EquipmentClassif4 = Utils.IsTrue(_cells.GetCell(57, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(57, i).Value)) : null,
                            EquipmentClassif5 = Utils.IsTrue(_cells.GetCell(58, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(58, i).Value)) : null,
                            EquipmentClassif6 = Utils.IsTrue(_cells.GetCell(59, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(59, i).Value)) : null,
                            EquipmentClassif7 = Utils.IsTrue(_cells.GetCell(60, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(60, i).Value)) : null,
                            EquipmentClassif8 = Utils.IsTrue(_cells.GetCell(61, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(61, i).Value)) : null,
                            EquipmentClassif9 = Utils.IsTrue(_cells.GetCell(62, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(62, i).Value)) : null,
                            EquipmentClassif10 = Utils.IsTrue(_cells.GetCell(63, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(63, i).Value)) : null,
                            EquipmentClassif11 = Utils.IsTrue(_cells.GetCell(64, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(64, i).Value)) : null,
                            EquipmentClassif12 = Utils.IsTrue(_cells.GetCell(65, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(65, i).Value)) : null,
                            EquipmentClassif13 = Utils.IsTrue(_cells.GetCell(66, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(66, i).Value)) : null,
                            EquipmentClassif14 = Utils.IsTrue(_cells.GetCell(67, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(67, i).Value)) : null,
                            EquipmentClassif15 = Utils.IsTrue(_cells.GetCell(68, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(68, i).Value)) : null,
                            EquipmentClassif16 = Utils.IsTrue(_cells.GetCell(69, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(69, i).Value)) : null,
                            EquipmentClassif17 = Utils.IsTrue(_cells.GetCell(70, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(70, i).Value)) : null,
                            EquipmentClassif18 = Utils.IsTrue(_cells.GetCell(71, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(71, i).Value)) : null,
                            EquipmentClassif19 = Utils.IsTrue(_cells.GetCell(72, validationRow).Value) ? Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(72, i).Value)) : null
                        }
                    };

                    EquipmentActions.UpdateEquipmentData(opSheet, urlService, equipment);

                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn01, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:UpdateEquipment()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void UpdateEquipmentStatus()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;

            var opSheet = new EquipmentService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = _eFunctions.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    
                    var equipmentRef = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);

                    List<string> equipmentList = EquipmentActions.GetEquipmentList(_eFunctions, opSheet.district, equipmentRef);
                    var equipment = new Equipment
                    {
                        EquipmentNo = equipmentList.Any() ? equipmentList.First() : equipmentRef,
                        EquipmentStatus = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(2, i).Value))
                    };
                    var newStatus = EquipmentActions.UpdateEquipmentStatus(opSheet, urlService, equipment.EquipmentNo, equipment.EquipmentStatus);

                    if(equipment.EquipmentStatus != null && !equipment.EquipmentStatus.Equals(newStatus))
                        throw new Exception("No se ha podido establecer el estado especificado " + equipment.EquipmentStatus + ". Estado actual " + newStatus);
                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn01, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:UpdateEquipment()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void DeleteEquipment()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;

            var opSheet = new EquipmentService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = _eFunctions.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

                    var equipmentRef = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);

                    var equipmentList = EquipmentActions.GetEquipmentList(_eFunctions, opSheet.district, equipmentRef);
                    var equipment = new Equipment
                    {
                        EquipmentNo = equipmentList.Any() ? equipmentList.First() : equipmentRef,
                        EquipmentStatus = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(2, i).Value))
                    };

                    throw new NotImplementedException();

                    _cells.GetCell(ResultColumn01, i).Value = "ELIMINADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn01, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:UpdateEquipment()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(ResultColumn01, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void ExecuteTraceAction()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);
            var i = TitleRow02 + 1;

            var opSheet = new EquipTraceService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = _eFunctions.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

                    var instEquipmentRef = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    var compCode = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var compModCode = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    var action = Utils.GetCodeKey(_cells.GetEmptyIfNull(_cells.GetCell(4, i).Value));
                    var fitEquipmentRef = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                    var fitDate = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value);
                    var refType = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value);
                    var refNumber = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value);
                    
                    List<string> instEquipmentList = EquipmentActions.GetEquipmentList(_eFunctions, opSheet.district, instEquipmentRef);
                    List<string> fitEquipmentList = EquipmentActions.GetEquipmentList(_eFunctions, opSheet.district, fitEquipmentRef);
                    
                    var instEquipmentNo = instEquipmentList.Any() ? instEquipmentList.First() : instEquipmentRef;
                    var fitEquipmentNo = fitEquipmentList.Any() ? fitEquipmentList.First() : fitEquipmentRef;

                    bool result;
                    if(action.ToUpper().Equals("F"))
                        result = TracingActions.Fitment(opSheet, urlService, instEquipmentNo, compCode, compModCode, fitEquipmentNo, fitDate, refType, refNumber);
                    else if (action.ToUpper().Equals("D"))
                        result = TracingActions.Defitment(opSheet, urlService, instEquipmentNo, compCode, compModCode, fitEquipmentNo, fitDate, refType, refNumber);
                    else
                        throw new Exception("No se ha seleccionado una acción a realizar");

                    var replyMessage = result ? "SE HA REALIZADO LA ACCION" : "HA OCURRIDO UN ERROR INESPERADO";
                    var styleResult = result ? StyleConstants.Success : StyleConstants.Error;

                    _cells.GetCell(ResultColumn02, i).Value = replyMessage;
                    _cells.GetCell(1, i).Style = styleResult;
                    _cells.GetCell(ResultColumn02, i).Style = styleResult;
                    _cells.GetCell(ResultColumn02, i).Select();
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn02, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:UpdateEquipment()", ex.Message, _eFunctions.DebugErrors);
                }
                finally
                {
                    _cells.GetCell(ResultColumn02, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
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

        
    }
}
