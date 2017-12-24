using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Constants;
using EllipseStandardJobsClassLibrary;
using StandardJobService = EllipseStandardJobsClassLibrary.StandardJobService;
using StandardJobTaskService = EllipseStandardJobsClassLibrary.StandardJobTaskService;
using ResourceReqmntsService = EllipseStandardJobsClassLibrary.ResourceReqmntsService;
using MaterialReqmntsService = EllipseStandardJobsClassLibrary.MaterialReqmntsService;
using System.Web.Services.Ellipse;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EquipmentReqmntsService = EllipseStandardJobsClassLibrary.EquipmentReqmntsService;
using EllipseStdTextClassLibrary;
// ReSharper disable UseNullPropagation

namespace EllipseStandardJobsExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Excel.Application _excelApp;
        private const string SheetName01 = "StandardJobs";
        private const string SheetName02 = "SJTasks";
        private const string SheetName03 = "SJRequirements";
        private const string SheetNameQualRev = "SJRevisionCalidad";
        private const string SheetName05 = "SJRefCodes";

        private const int TitleRow01 = 8;
        private const int TitleRow02 = 6;
        private const int TitleRow03 = 6;
        private const int TitleRowQualRev = 8;
        private const int TitleRow05 = 7;
        private const int ResultColumn01 = 40;
        private const int ResultColumn02 = 26;
        private const int ResultColumn03 = 14;
        private const int ResultColumnQualRev = 26;
        private const int ResultColumn05 = 35;
        private const string TableName01 = "StandardJobsTable";
        private const string TableName02 = "SJTasksTable";
        private const string TableName03 = "SJRequirementsTable";
        private const string TableNameQualRev = "SJRevCalTable";
        private const string TableName05 = "SJJobCodesTable";
        private bool ActiveQualityValidation = true;
        private bool _quickReview = true;
        private string _standardStatus = "A";//A Active / I Inactive
        private Thread _thread;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var enviroments = Environments.GetEnviromentList();
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

        private void btnStandardReview_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _quickReview = false;
                _thread = new Thread(ReviewStandardJobs);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnQuickStandardReview_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _quickReview = true;
                _thread = new Thread(ReviewStandardJobs);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnReReviewStandard_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(ReReviewStandardJobs);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnCreateStandard_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(CreateStandardJobList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnModifyStandard_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(UpdateStandardJobList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnActivateStandard_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _standardStatus = "A";
                _thread = new Thread(UpdateStandardJobStatus);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnDeactivateStandard_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _standardStatus = "I";
                _thread = new Thread(UpdateStandardJobStatus);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }


        private void btnCleanStandardTable_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName01);
        }

        private void btnReviewTasks_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                _thread = new Thread(ReviewStandardJobTasks);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnExecuteTaskActions_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(ExecuteTaskActionsPost);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnCleanTasksTable_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName02);
        }

        private void btnReviewQualityStdJobs_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameQualRev)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(QualityReviewStandardJobs);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnUpdateQualityStdJobs_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetNameQualRev)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(UpdateQualityStandardList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnCleanQualityStdJobsTable_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableNameQualRev);
        }

        private void btnReviewRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewRequirements);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnExecuteRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                _thread = new Thread(ExecuteRequirementActions);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnCleanRequirementTable_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.ClearTableRange(TableName03);
        }

        private void btnGetAplRequirements_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName03)
            {

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                _thread = new Thread(GetAplTaskRequirements);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;

                #region CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 4)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                if (ActiveQualityValidation)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(1).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STANDARD JOBS - ELLIPSE 8";
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

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), false);
                _cells.SetValidationList(_cells.GetCell("B4"),
                    Groups.GetWorkGroupList().Select(wg => wg.Name).ToList());

                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow01 - 1)
                    .AddComment("Solo se modificará este campo si el valor es verdadero (VERDADERO, TRUE, Y, 1)");
                _cells.GetCell(5, TitleRow01 - 1).Value = "true";
                _cells.GetRange(3, TitleRow01 - 1, ResultColumn01 - 1, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;

                for (var i = 10; i < 25; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1)
                        .AddComment("Solo se modificará este campo si el valor es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }

                //GENERAL
                _cells.GetRange(1, TitleRow01 - 2, 5, TitleRow01 - 2).Style = StyleConstants.Option;
                _cells.GetCell(1, TitleRow01 - 2).Value = "GENERAL";
                _cells.GetRange(1, TitleRow01 - 2, 5, TitleRow01 - 2).Merge();

                _cells.GetCell(1, TitleRow01).Value = "DISTRICT";
                _cells.GetCell(2, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(3, TitleRow01).Value = "STD_JOB_NO";
                _cells.GetCell(3, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(4, TitleRow01).Value = "SJ_STATUS";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetCell(5, TitleRow01).Value = "STD_JOB_DESC";

                //CONSULTA
                _cells.GetRange(6, TitleRow01 - 2, 9, TitleRow01 - 2).Style = StyleConstants.Option;
                _cells.GetCell(6, TitleRow01 - 2).Value = "CONSULTA";
                _cells.GetRange(6, TitleRow01 - 2, 9, TitleRow01 - 2).Merge();

                _cells.GetCell(6, TitleRow01).Value = "USO_OTS";
                _cells.GetCell(7, TitleRow01).Value = "USO_MSTS";
                _cells.GetCell(8, TitleRow01).Value = "ULTIMO_USO";
                _cells.GetCell(9, TitleRow01).Value = "NO_OF_TASKS";
                _cells.GetRange(6, TitleRow01, 9, TitleRow01).Style = StyleConstants.TitleInformation;

                //PLANNING
                _cells.GetRange(10, TitleRow01 - 2, 20, TitleRow01 - 2).Style = StyleConstants.Option;
                _cells.GetCell(10, TitleRow01 - 2).Value = "PLANNING";
                _cells.GetRange(10, TitleRow01 - 2, 20, TitleRow01 - 2).Merge();

                _cells.GetCell(10, TitleRow01).Value = "ORIGINATOR_ID";
                _cells.GetCell(10, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(11, TitleRow01).Value = "ASSIGN_PERSON";
                _cells.GetCell(12, TitleRow01).Value = "ORIG_PRIORITY";
                var woTypeCodes = WoTypeMtType.GetWoTypeList();
                var mtTypeCodes = WoTypeMtType.GetMtTypeList();

                _cells.GetCell(13, TitleRow01).Value = "WO_TYPE";
                _cells.SetValidationList(_cells.GetCell(13, TitleRow01 + 1), new List<string>(woTypeCodes.Keys));
                _cells.GetCell(14, TitleRow01).Value = "MT_TYPE";
                _cells.SetValidationList(_cells.GetCell(14, TitleRow01 + 1), new List<string>(mtTypeCodes.Keys));

                _cells.GetCell(15, TitleRow01).Value = "COMP_CODE";
                _cells.GetCell(16, TitleRow01).Value = "COMP_MOD_CODE";
                _cells.GetCell(17, TitleRow01).Value = "UNITS_OF_WORK";
                _cells.GetCell(18, TitleRow01).Value = "UNITS_REQUIRED";
                _cells.GetCell(18, TitleRow01).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(19, TitleRow01).Value = "EST_DUR_HRS_FLAG";
                _cells.GetCell(19, TitleRow01).AddComment("Y/N Y: Usar las horas de duración como calculadas; N: Usa las horas de duración estimadas");
                _cells.GetCell(20, TitleRow01).Value = "EST_DUR_HRS";
                _cells.GetRange(15, TitleRow01, 20, TitleRow01).Style = StyleConstants.TitleOptional;

                //COSTS
                _cells.GetRange(21, TitleRow01 - 2, 28, TitleRow01 - 2).Style = StyleConstants.Option;
                _cells.GetCell(21, TitleRow01 - 2).Value = "COSTS";
                _cells.GetRange(21, TitleRow01 - 2, 28, TitleRow01 - 2).Merge();

                _cells.GetCell(21, TitleRow01).Value = "ACCOUNT_CODE";

                _cells.GetCell(22, TitleRow01).Value = "REALL_ACCT_CDE";
                _cells.GetCell(22, TitleRow01).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(23, TitleRow01).Value = "PROJECT_NO";
                _cells.GetCell(23, TitleRow01).NumberFormat = NumberFormatConstants.Text;
                _cells.GetCell(24, TitleRow01).Value = "EST_OTH_COST";
                _cells.GetRange(22, TitleRow01, 24, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(25, TitleRow01).Value = "CALC_LAB_HRS";
                _cells.GetCell(26, TitleRow01).Value = "CALC_LAB_COST";
                _cells.GetCell(27, TitleRow01).Value = "CALC_MAT_COST";
                _cells.GetCell(28, TitleRow01).Value = "CALC_EQUIP_COST";
                _cells.GetRange(25, TitleRow01, 28, TitleRow01).Style = StyleConstants.TitleInformation;
                //JOB CODES
                _cells.GetRange(29, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2).Style = StyleConstants.Option;
                _cells.GetCell(29, TitleRow01 - 2).Value = "JOB CODES";
                _cells.GetRange(29, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2).Merge();

                for (var i = 29; i < ResultColumn01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1)
                        .AddComment("Solo se modificará este campo si el valor es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }

                _cells.GetCell(29, TitleRow01).Value = "JOB CODE 01";
                _cells.GetCell(30, TitleRow01).Value = "JOB CODE 02";
                _cells.GetCell(31, TitleRow01).Value = "JOB CODE 03";
                _cells.GetCell(32, TitleRow01).Value = "JOB CODE 04";
                _cells.GetCell(33, TitleRow01).Value = "JOB CODE 05";
                _cells.GetCell(34, TitleRow01).Value = "JOB CODE 06";
                _cells.GetCell(35, TitleRow01).Value = "JOB CODE 07";
                _cells.GetCell(36, TitleRow01).Value = "JOB CODE 08";
                _cells.GetCell(37, TitleRow01).Value = "JOB CODE 09";
                _cells.GetCell(38, TitleRow01).Value = "JOB CODE 10";
                _cells.GetRange(29, TitleRow01, 38, TitleRow01).Style = StyleConstants.TitleOptional;

                _cells.GetCell(39, TitleRow01).Value = "DESCRIPCION EXTENDIDA";
                _cells.GetCell(39, TitleRow01).Style = StyleConstants.TitleOptional;

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 2
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STD JOBS TASKS - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetRange(1, TitleRow02, ResultColumn02 - 1, TitleRow02).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, TitleRow02 - 1).Value = "STANDARD";
                _cells.GetRange(1, TitleRow02 - 1, 5, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, TitleRow02 - 1, 5, TitleRow02 - 1).Merge();

                _cells.GetCell(1, TitleRow02).Value = "DISTRICT";
                _cells.GetCell(2, TitleRow02).Value = "WORK_GROUP";
                _cells.GetCell(3, TitleRow02).Value = "STD_JOB_NO";
                _cells.GetCell(4, TitleRow02).Value = "STD_JOB_DESC";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleInformation;
                //ACTION
                _cells.GetCell(5, TitleRow02).Value = "ACTION";
                _cells.GetCell(5, TitleRow02).Style = StyleConstants.TitleAction;
                _cells.GetCell(5, TitleRow02).AddComment("C: Crear \nM: Modificar \nD: Eliminar");
                _cells.SetValidationList(_cells.GetCell(5, TitleRow02 + 1), new List<string> { "C", "M", "D" });
                //GENERAL
                _cells.GetCell(6, TitleRow02 - 1).Value = "GENERAL";
                _cells.GetRange(6, TitleRow02 - 1, 11, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(6, TitleRow02 - 1, 11, TitleRow02 - 1).Merge();

                _cells.GetCell(6, TitleRow02).Value = "TASK_NO";
                _cells.GetCell(7, TitleRow02).Value = "SJ_TASK_DESC";
                _cells.GetCell(8, TitleRow02).Value = "JOB_DESC_CODE";
                _cells.GetCell(9, TitleRow02).Value = "SAFETY_INST";
                _cells.GetCell(10, TitleRow02).Value = "COMPL_INST";
                _cells.GetCell(11, TitleRow02).Value = "COMPL_TEXT_CODE";

                //PLANNING
                _cells.GetCell(12, TitleRow02 - 1).Value = "PLANNING";
                _cells.GetRange(12, TitleRow02 - 1, 16, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(12, TitleRow02 - 1, 16, TitleRow02 - 1).Merge();

                _cells.GetCell(12, TitleRow02).Value = "ASSIGN_PERSON";
                _cells.GetCell(13, TitleRow02).Value = "EST_MACH_HRS";
                _cells.GetCell(14, TitleRow02).Value = "UNIT_OF_WORK";
                _cells.GetCell(15, TitleRow02).Value = "UNITS_REQUIRED";
                _cells.GetCell(16, TitleRow02).Value = "UNITS_PER_DAY";
                _cells.GetRange(12, TitleRow02, 16, TitleRow02).Style = StyleConstants.TitleOptional;

                //RECURSOS
                _cells.GetCell(17, TitleRow02 - 1).Value = "RECURSOS";
                _cells.GetRange(17, TitleRow02 - 1, 19, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(17, TitleRow02 - 1, 19, TitleRow02 - 1).Merge();

                _cells.GetCell(17, TitleRow02).Value = "EST_DUR_HRS";
                _cells.GetCell(17, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(18, TitleRow02).Value = "LABOR";
                _cells.GetCell(19, TitleRow02).Value = "MATERIAL";
                _cells.GetRange(18, TitleRow02, 19, TitleRow02).Style = StyleConstants.TitleInformation;

                //APL
                _cells.GetCell(20, TitleRow02 - 1).Value = "APL";
                _cells.GetRange(20, TitleRow02 - 1, 24, TitleRow02 - 1).Style = StyleConstants.Option;
                _cells.GetRange(20, TitleRow02 - 1, 24, TitleRow02 - 1).Merge();

                _cells.GetCell(20, TitleRow02).Value = "EQUIP_GRP_ID";
                _cells.GetCell(21, TitleRow02).Value = "APL_TYPE";
                _cells.GetCell(22, TitleRow02).Value = "COMP_CODE";
                _cells.GetCell(23, TitleRow02).Value = "COMP_MOD_CODE";
                _cells.GetCell(24, TitleRow02).Value = "APL_SEQ_NO";

                _cells.GetRange(20, TitleRow02, 24, TitleRow02).Style = StyleConstants.TitleOptional;


                _cells.GetCell(25, TitleRow02).Value = "DESCRIPCION EXTENDIDA";
                _cells.GetCell(25, TitleRow02).Style = StyleConstants.TitleOptional;
                //RESULTADO
                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA 3

                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(3).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName03;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STDJ TASK REQUIREMENTS - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetRange(1, TitleRow03, ResultColumn03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, TitleRow03 - 1).Value = "STANDARD / TASK";
                _cells.GetRange(1, TitleRow03 - 1, 6, TitleRow03 - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, TitleRow03 - 1, 6, TitleRow03 - 1).Merge();

                _cells.GetCell(1, TitleRow03).Value = "DISTRICT";
                _cells.GetCell(2, TitleRow03).Value = "WORK_GROUP";
                _cells.GetCell(3, TitleRow03).Value = "STD_JOB_NO";
                _cells.GetCell(4, TitleRow03).Value = "TASK_NO";
                _cells.GetCell(5, TitleRow03).Value = "SJ_TASK_DESC";

                //ACTION
                _cells.GetCell(6, TitleRow03).Value = "ACTION";
                _cells.GetCell(6, TitleRow03).Style = StyleConstants.TitleAction;
                _cells.GetCell(6, TitleRow03).AddComment("C: Crear Requerimiento \nM: Modificar Requerimiento \nD: Eliminar Requerimiento");
                _cells.SetValidationList(_cells.GetCell(6, TitleRow03 + 1), new List<string> { "C", "M", "D" });
                //GENERAL
                _cells.GetCell(7, TitleRow03 - 1).Value = "GENERAL";
                _cells.GetRange(7, TitleRow03 - 1, 13, TitleRow03 - 1).Style = StyleConstants.Option;
                _cells.GetRange(7, TitleRow03 - 1, 13, TitleRow03 - 1).Merge();

                _cells.GetCell(7, TitleRow03).Value = "REQ_TYPE";
                _cells.GetCell(7, TitleRow03).AddComment("LAB: LABOR\nMAT: MATERIAL");
                _cells.SetValidationList(_cells.GetCell(7, TitleRow03 + 1), new List<string> { "LAB", "MAT" });
                _cells.GetCell(8, TitleRow03).Value = "SEQ_NO";
                _cells.GetCell(8, TitleRow03).AddComment("Aplica solo para Creación y Modificación de Requerimientos");
                _cells.GetCell(9, TitleRow03).Value = "REQ_CODE";
                _cells.GetCell(9, TitleRow03).AddComment("Recurso: Class+Code (Ver hoja de recursos) \nMaterial: StockCode");
                _cells.GetCell(10, TitleRow03).Value = "DESCRIPTION";
                _cells.GetCell(11, TitleRow03).Value = "QTY REQ";
                _cells.GetCell(12, TitleRow03).Value = "HRS_REQ";
                _cells.GetCell(12, TitleRow03).AddComment("Horas requeridas del recurso. (Solo aplica para labor)");
                _cells.GetCell(13, TitleRow03).Value = "UOM";
                _cells.GetCell(13, TitleRow03).AddComment("Unidad de Medida del Recurso. (Solo aplica para Equipos)");


                //RESULTADO
                _cells.GetCell(ResultColumn03, TitleRow03).Value = "RESULTADO";
                _cells.GetCell(ResultColumn03, TitleRow03).Style = StyleConstants.TitleResult;

                _cells.GetRange(1, TitleRow03 + 1, ResultColumn03 - 2, TitleRow03 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow03, ResultColumn03, TitleRow03 + 1), TableName03);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                #endregion

                #region CONSTRUYO LA HOJA QualRev DE REVISIÓN DE CALIDAD

                // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                if (ActiveQualityValidation)
                {
                    // ReSharper disable once UseIndexedProperty
                    _excelApp.ActiveWorkbook.Sheets.get_Item(4).Activate();
                    _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetNameQualRev;

                    _cells.GetCell("A1").Value = "CERREJÓN";
                    _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                    _cells.MergeCells("A1", "B2");

                    _cells.GetCell("C1").Value = "STANDARD JOBS - ELLIPSE 8";
                    _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                    _cells.MergeCells("C1", "J2");

                    _cells.GetCell("K1").Value = "OBLIGATORIO";
                    _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                    _cells.GetCell("K2").Value = "OPCIONAL";
                    _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                    _cells.GetCell("K3").Value = "INFORMATIVO";
                    _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                    _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                    _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                    _cells.GetCell("A3").Value = "DISTRITO";
                    _cells.GetCell("B3").Value = "ICOR";
                    _cells.GetCell("A4").Value = "GRUPO";
                    _cells.GetRange("A3", "A4").Style = StyleConstants.Option;
                    _cells.GetRange("B3", "B4").Style = StyleConstants.Select;

                    //Adicionar validaciones
                    _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList());
                    _cells.SetValidationList(_cells.GetCell("B4"),
                        Groups.GetWorkGroupList().Select(wg => wg.Name).ToList());

                    _cells.GetRange(1, TitleRowQualRev, ResultColumnQualRev, TitleRowQualRev).Style =
                        StyleConstants.TitleRequired;

                    for (var i = 6; i < ResultColumnQualRev; i++)
                    {
                        _cells.GetCell(i, TitleRowQualRev - 1)
                            .AddComment("Solo se modificará este campo si el valor es verdadero (VERDADERO, TRUE, Y, 1)");
                        _cells.GetCell(i, TitleRowQualRev - 1).Value = "true";
                        _cells.GetCell(i, TitleRowQualRev - 1).Style = StyleConstants.ItalicSmall;
                    }

                    //GENERAL
                    _cells.GetRange(1, TitleRowQualRev - 2, 5, TitleRowQualRev - 2).Style = StyleConstants.Option;
                    _cells.GetCell(1, TitleRowQualRev - 2).Value = "GENERAL";
                    _cells.GetRange(1, TitleRowQualRev - 2, 5, TitleRowQualRev - 2).Merge();

                    _cells.GetCell(1, TitleRowQualRev).Value = "DISTRICT";
                    _cells.GetCell(2, TitleRowQualRev).Value = "WORK_GROUP";
                    _cells.GetCell(3, TitleRowQualRev).Value = "STD_JOB_NO";
                    _cells.GetCell(3, TitleRowQualRev + 1).NumberFormat = NumberFormatConstants.Text;
                    _cells.GetCell(4, TitleRowQualRev).Value = "SJ_STATUS";
                    _cells.GetCell(4, TitleRowQualRev).Style = StyleConstants.TitleRequired;
                    _cells.GetCell(5, TitleRowQualRev).Value = "STD_JOB_DESC";
                    _cells.GetCell(5, TitleRowQualRev).Style = StyleConstants.TitleInformation;

                    //CONSULTA
                    _cells.GetCell(6, TitleRowQualRev).Value = "ORIG_PRIORITY";
                    _cells.GetCell(7, TitleRowQualRev).Value = "WO_TYPE";
                    _cells.SetValidationList(_cells.GetCell(7, TitleRowQualRev + 1),
                        new List<string>(woTypeCodes.Keys));
                    _cells.GetCell(8, TitleRowQualRev).Value = "MT_TYPE";
                    _cells.SetValidationList(_cells.GetCell(8, TitleRowQualRev + 1),
                        new List<string>(mtTypeCodes.Keys));
                    _cells.GetCell(9, TitleRowQualRev).Value = "UNITS_OF_WORK";
                    _cells.GetCell(10, TitleRowQualRev).Value = "UNITS_REQUIRED";
                    _cells.GetCell(10, TitleRowQualRev).NumberFormat = NumberFormatConstants.Text;
                    _cells.GetCell(11, TitleRowQualRev).Value = "EST_DUR_HRS_FLAG";
                    _cells.GetCell(11, TitleRowQualRev).AddComment("Y/N Y: Usar las horas de duración como calculadas; N: Usa las horas de duración estimadas");
                    _cells.GetCell(12, TitleRowQualRev).Value = "CALC_RES_FLAG";
                    _cells.GetCell(12, TitleRowQualRev).AddComment("Y/N Y: Usar las horas de duración como calculadas; N: Usa las horas de duración estimadas");
                    _cells.GetCell(13, TitleRowQualRev).Value = "CALC_MAT_FLAG";
                    _cells.GetCell(13, TitleRowQualRev).AddComment("Y/N Y: Usar las horas de duración como calculadas; N: Usa las horas de duración estimadas");
                    _cells.GetCell(14, TitleRowQualRev).Value = "CALC_EQU_FLAG";
                    _cells.GetCell(14, TitleRowQualRev).AddComment("Y/N Y: Usar las horas de duración como calculadas; N: Usa las horas de duración estimadas");
                    _cells.GetCell(15, TitleRowQualRev).Value = "CALC_TOT_FLAG";
                    _cells.GetCell(15, TitleRowQualRev).AddComment("Y/N Y: Usar las horas de duración como calculadas; N: Usa las horas de duración estimadas");
                    _cells.GetCell(16, TitleRowQualRev).Value = "JOB_CODE1";
                    _cells.GetCell(17, TitleRowQualRev).Value = "JOB_CODE2";
                    _cells.GetCell(18, TitleRowQualRev).Value = "JOB_CODE3";
                    _cells.GetCell(19, TitleRowQualRev).Value = "JOB_CODE4";
                    _cells.GetCell(20, TitleRowQualRev).Value = "JOB_CODE5";
                    _cells.GetCell(21, TitleRowQualRev).Value = "JOB_CODE6";
                    _cells.GetCell(22, TitleRowQualRev).Value = "JOB_CODE7";
                    _cells.GetCell(23, TitleRowQualRev).Value = "JOB_CODE8";
                    _cells.GetCell(24, TitleRowQualRev).Value = "JOB_CODE9";
                    _cells.GetCell(25, TitleRowQualRev).Value = "JOB_CODE10";

                    _cells.GetCell(26, TitleRowQualRev).Value = "DESCRIPCION EXTENDIDA";
                    _cells.GetCell(26, TitleRowQualRev).Style = StyleConstants.TitleOptional;


                    _cells.GetCell(ResultColumnQualRev, TitleRowQualRev).Value = "RESULTADO";
                    _cells.GetCell(ResultColumnQualRev, TitleRowQualRev).Style = StyleConstants.TitleResult;

                    _cells.FormatAsTable(_cells.GetRange(1, TitleRowQualRev, ResultColumnQualRev, TitleRowQualRev + 1),
                        TableNameQualRev);
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }
                // 
                #endregion
                //CONSTRUYO LA HOJA 5 - REF CODE
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(5).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName05;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "STANDARD JOBS REFERENCE CODES - ELLIPSE 8";
                _cells.GetCell("C1").Style = StyleConstants.HeaderDefault;
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = StyleConstants.TitleRequired;
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = StyleConstants.TitleOptional;
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = StyleConstants.TitleInformation;
                _cells.GetCell("K4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K4").Style = StyleConstants.TitleAction;

                _cells.GetRange(1, TitleRow03, ResultColumn03 - 1, TitleRow03).Style = StyleConstants.TitleRequired;

                //STANDARD
                _cells.GetCell(1, TitleRow05 - 1).Value = "STANDARD";
                _cells.GetRange(1, TitleRow05 - 1, 4, TitleRow03 - 1).Style = StyleConstants.Option;
                _cells.GetRange(1, TitleRow05 - 1, 4, TitleRow03 - 1).Merge();

                _cells.GetCell(1, TitleRow05).Value = "DISTRICT";
                _cells.GetCell(2, TitleRow05).Value = "WORK_GROUP";
                _cells.GetCell(3, TitleRow05).Value = "STD_JOB_NO";
                _cells.GetCell(4, TitleRow05).Value = "STD_JOB_DESC";

                //JOB CODES
                _cells.GetCell(5, TitleRow05 - 2).Value = "REFERENCE CODES";
                _cells.GetRange(5, TitleRow05 - 2, 35, TitleRow05 - 2).Style = StyleConstants.Option;
                _cells.GetRange(5, TitleRow05 - 2, 35, TitleRow05 - 2).Merge();

                _cells.GetRange(1, TitleRow05, 3, TitleRow05).Style = StyleConstants.TitleRequired;
                _cells.GetRange(3, TitleRow05, ResultColumn05 - 1, TitleRow05).Style = StyleConstants.TitleOptional;

                for (var i = 5; i < ResultColumn05; i++)
                {
                    _cells.GetCell(i, TitleRow05 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow05 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow05 - 1).Value = "true";
                }

                _cells.GetCell(5, TitleRow05).Value = "Work Request";
                _cells.GetCell(6, TitleRow05).Value = "Comentarios Duraciones";
                _cells.GetCell(7, TitleRow05).Value = "Com.Dur. Text";
                _cells.GetCell(8, TitleRow05).Value = "EmpleadoId";
                _cells.GetCell(9, TitleRow05).Value = "Nro. Componente";
                _cells.GetCell(10, TitleRow05).Value = "P1. Eq.Liv-Med";
                _cells.GetCell(11, TitleRow05).Value = "P2. Eq.Movil-Minero";
                _cells.GetCell(12, TitleRow05).Value = "P3. Manejo Sust.Peligrosa";
                _cells.GetCell(13, TitleRow05).Value = "P4. Guardas Equipo";
                _cells.GetCell(14, TitleRow05).Value = "P5. Aislamiento";
                _cells.GetCell(15, TitleRow05).Value = "P6. Trabajos Altura";
                _cells.GetCell(16, TitleRow05).Value = "P7. Manejo Cargas";
                _cells.GetCell(17, TitleRow05).Value = "Proyecto ICN";
                _cells.GetCell(18, TitleRow05).Value = "Reembolsable";
                _cells.GetCell(19, TitleRow05).Value = "Fecha No Conforme";
                _cells.GetCell(20, TitleRow05).Value = "Fecha NC Text";
                _cells.GetCell(21, TitleRow05).Value = "No Conforme?";
                _cells.GetCell(22, TitleRow05).Value = "Fecha Ejecución";
                _cells.GetCell(23, TitleRow05).Value = "Hora Ingreso";
                _cells.GetCell(24, TitleRow05).Value = "Hora Salida";
                _cells.GetCell(25, TitleRow05).Value = "Nombre Buque";
                _cells.GetCell(26, TitleRow05).Value = "Calif. Encuesta";
                _cells.GetCell(27, TitleRow05).Value = "Tarea Crítica?";
                _cells.GetCell(28, TitleRow05).Value = "Garantía";
                _cells.GetCell(29, TitleRow05).Value = "Garantía Text";
                _cells.GetCell(30, TitleRow05).Value = "Cód. Certificación";
                _cells.GetCell(31, TitleRow05).Value = "Fecha Entrega";
                _cells.GetCell(32, TitleRow05).Value = "Relacionar EV";
                _cells.GetCell(33, TitleRow05).Value = "Departamento";
                _cells.GetCell(34, TitleRow05).Value = "Localización";

                _cells.GetCell(ResultColumn05, TitleRow05).Value = "RESULTADO";
                _cells.GetCell(ResultColumn05, TitleRow05).Style = StyleConstants.TitleResult;
                _cells.GetRange(1, TitleRow05 + 1, ResultColumn05, TitleRow05 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow05, ResultColumn05, TitleRow05 + 1), TableName05);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();


                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void ReviewStandardJobs()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(TableName01);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var stOpContext = StdText.GetStdTextOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, true);

            var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value2);
            var workGroup = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value2);

            var i = TitleRow01 + 1;

            var listJobs = StandardJobActions.FetchStandardJob(_eFunctions, districtCode, workGroup, _quickReview);

            foreach (StandardJob stdJob in listJobs)
            {
                try
                {
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "" + stdJob.DistrictCode;
                    _cells.GetCell(2, i).Value = "" + stdJob.WorkGroup;
                    _cells.GetCell(3, i).Value = "" + stdJob.StandardJobNo;
                    _cells.GetCell(4, i).Value = "" + stdJob.Status;
                    _cells.GetCell(5, i).Value = "" + stdJob.StandardJobDescription;
                    //CONSULTA                   
                    _cells.GetCell(6, i).Value = "" + stdJob.NoWos;
                    _cells.GetCell(7, i).Value = "" + stdJob.NoWos;
                    _cells.GetCell(8, i).Value = "" + stdJob.LastUse;
                    _cells.GetCell(9, i).Value = "" + stdJob.NoTasks;
                    //PLANNING
                    _cells.GetCell(10, i).Value = "" + stdJob.OriginatorId;
                    _cells.GetCell(11, i).Value = "" + stdJob.AssignPerson;
                    _cells.GetCell(12, i).Value = "" + stdJob.OrigPriority;
                    _cells.GetCell(12, i).Style = !WoTypeMtType.ValidatePriority(stdJob.OrigPriority)
                        ? StyleConstants.Error : StyleConstants.Normal;
                    _cells.GetCell(13, i).Value = "" + stdJob.WorkOrderType;
                    _cells.GetCell(14, i).Value = "" + stdJob.MaintenanceType;
                    _cells.GetRange(13, i, 14, i).Style = !WoTypeMtType.ValidateWoMtTypeCode(stdJob.WorkOrderType, stdJob.MaintenanceType)
                        ? StyleConstants.Error : StyleConstants.Normal;
                    _cells.GetCell(15, i).Value = "" + stdJob.CompCode;
                    _cells.GetCell(16, i).Value = "" + stdJob.CompModCode;
                    _cells.GetCell(17, i).Value = "" + stdJob.UnitOfWork;
                    _cells.GetCell(18, i).Value = "" + stdJob.UnitsRequired;
                    if (!string.IsNullOrWhiteSpace(stdJob.UnitOfWork))
                        _cells.GetRange(17, i, 18, i).Style = int.Parse(stdJob.UnitsRequired) > 0
                            ? StyleConstants.Warning : StyleConstants.Error;
                    else
                        _cells.GetRange(17, i, 18, i).Style = StyleConstants.Normal;

                    _cells.GetCell(19, i).Value = "" + stdJob.CalculatedDurationsHrsFlg;
                    _cells.GetCell(20, i).Value = "" + stdJob.EstimatedDurationsHrs;
                    //COSTS
                    _cells.GetCell(21, i).Value = "" + stdJob.AccountCode;
                    _cells.GetCell(22, i).Value = "" + stdJob.ReallocAccCode;
                    _cells.GetCell(23, i).Value = "" + stdJob.ProjectNo;

                    _cells.GetCell(24, i).Value = "" + stdJob.EstimatedOtherCost;
                    _cells.GetCell(25, i).Value = "" + stdJob.CalculatedLabHrs;
                    _cells.GetCell(26, i).Value = "" + stdJob.CalculatedLabCost;
                    _cells.GetCell(27, i).Value = "" + stdJob.CalculatedMatCost;
                    _cells.GetCell(28, i).Value = "" + stdJob.CalculatedEquipmentCost;
                    //JOB CODES
                    _cells.GetCell(29, i).Value = "" + stdJob.JobCode1;
                    _cells.GetCell(30, i).Value = "" + stdJob.JobCode2;
                    _cells.GetCell(31, i).Value = "" + stdJob.JobCode3;
                    _cells.GetCell(32, i).Value = "" + stdJob.JobCode4;
                    _cells.GetCell(33, i).Value = "" + stdJob.JobCode5;
                    _cells.GetCell(34, i).Value = "" + stdJob.JobCode6;
                    _cells.GetCell(35, i).Value = "" + stdJob.JobCode7;
                    _cells.GetCell(36, i).Value = "" + stdJob.JobCode8;
                    _cells.GetCell(37, i).Value = "" + stdJob.JobCode9;
                    _cells.GetCell(38, i).Value = "" + stdJob.JobCode10;
                    
                    var stdTextId = "SJ" + stdJob.DistrictCode + stdJob.StandardJobNo;
                    var extendedDescription = StdText.GetText(urlService, stOpContext, stdTextId);
                    _cells.GetCell(39, i).Value = extendedDescription;
                    _cells.GetCell(39, i).WrapText = false;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReviewStandardJobs()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(3, i).Select();
                    i++;
                }
            }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }


        private void ReReviewStandardJobs()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var stOpContext = StdText.GetStdTextOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, true);

            var i = TitleRow01 + 1;
            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                    var workGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2);
                    var stdJobNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2);

                    var stdJob = StandardJobActions.FetchStandardJob(_eFunctions, districtCode, workGroup, stdJobNo);
                    if (stdJob != null)
                    {
                        //Para resetear el estilo
                        _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + stdJob.DistrictCode;
                        _cells.GetCell(2, i).Value = "" + stdJob.WorkGroup;
                        _cells.GetCell(3, i).Value = "" + stdJob.StandardJobNo;
                        _cells.GetCell(4, i).Value = "" + stdJob.Status;
                        _cells.GetCell(5, i).Value = "" + stdJob.StandardJobDescription;
                        //CONSULTA                   
                        _cells.GetCell(6, i).Value = "" + stdJob.NoWos;
                        _cells.GetCell(7, i).Value = "" + stdJob.NoWos;
                        _cells.GetCell(8, i).Value = "" + stdJob.LastUse;
                        _cells.GetCell(9, i).Value = "" + stdJob.NoTasks;
                        //PLANNING
                        _cells.GetCell(10, i).Value = "" + stdJob.OriginatorId;
                        _cells.GetCell(11, i).Value = "" + stdJob.AssignPerson;
                        _cells.GetCell(12, i).Value = "" + stdJob.OrigPriority;
                        _cells.GetCell(12, i).Style = !WoTypeMtType.ValidatePriority(stdJob.OrigPriority)
                            ? StyleConstants.Error : StyleConstants.Normal;
                        _cells.GetCell(13, i).Value = "" + stdJob.WorkOrderType;
                        _cells.GetCell(14, i).Value = "" + stdJob.MaintenanceType;
                        _cells.GetRange(13, i, 14, i).Style = !WoTypeMtType.ValidateWoMtTypeCode(stdJob.WorkOrderType, stdJob.MaintenanceType)
                            ? StyleConstants.Error : StyleConstants.Normal;
                        _cells.GetCell(15, i).Value = "" + stdJob.CompCode;
                        _cells.GetCell(16, i).Value = "" + stdJob.CompModCode;
                        _cells.GetCell(17, i).Value = "" + stdJob.UnitOfWork;
                        _cells.GetCell(18, i).Value = "" + stdJob.UnitsRequired;
                        if (!string.IsNullOrWhiteSpace(stdJob.UnitOfWork))
                            _cells.GetRange(17, i, 18, i).Style = int.Parse(stdJob.UnitsRequired) > 0
                                ? StyleConstants.Warning : StyleConstants.Error;
                        else
                            _cells.GetRange(17, i, 18, i).Style = StyleConstants.Normal;
                        _cells.GetCell(19, i).Value = "" + stdJob.CalculatedDurationsHrsFlg;
                        _cells.GetCell(20, i).Value = "" + stdJob.EstimatedDurationsHrs;
                        //COSTS
                        _cells.GetCell(21, i).Value = "" + stdJob.AccountCode;
                        _cells.GetCell(22, i).Value = "" + stdJob.ReallocAccCode;
                        _cells.GetCell(23, i).Value = "" + stdJob.ProjectNo;

                        _cells.GetCell(24, i).Value = "" + stdJob.EstimatedOtherCost;
                        _cells.GetCell(25, i).Value = "" + stdJob.CalculatedLabHrs;
                        _cells.GetCell(26, i).Value = "" + stdJob.CalculatedLabCost;
                        _cells.GetCell(27, i).Value = "" + stdJob.CalculatedMatCost;
                        _cells.GetCell(28, i).Value = "" + stdJob.CalculatedEquipmentCost;
                        //JOB CODES
                        _cells.GetCell(29, i).Value = "" + stdJob.JobCode1;
                        _cells.GetCell(30, i).Value = "" + stdJob.JobCode2;
                        _cells.GetCell(31, i).Value = "" + stdJob.JobCode3;
                        _cells.GetCell(32, i).Value = "" + stdJob.JobCode4;
                        _cells.GetCell(33, i).Value = "" + stdJob.JobCode5;
                        _cells.GetCell(34, i).Value = "" + stdJob.JobCode6;
                        _cells.GetCell(35, i).Value = "" + stdJob.JobCode7;
                        _cells.GetCell(36, i).Value = "" + stdJob.JobCode8;
                        _cells.GetCell(37, i).Value = "" + stdJob.JobCode9;
                        _cells.GetCell(38, i).Value = "" + stdJob.JobCode10;
                        var stdTextId = "SJ" + stdJob.DistrictCode + stdJob.StandardJobNo;
                        var extendedDescription = StdText.GetText(urlService, stOpContext, stdTextId);
                        _cells.GetCell(39, i).Value = extendedDescription;
                        _cells.GetCell(39, i).WrapText = false;
                    }
                    else
                    {
                        for (var j = 4; j < ResultColumn01; j++)
                            _cells.GetCell(j, i).Value2 = "";
                    }

                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ReReviewStandardJobs()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(3, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ReviewStandardJobTasks()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(TableName02);
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            var stOpContext = StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, true);

            var stdCells = new ExcelStyleCells(_excelApp, SheetName01);
            stdCells.SetAlwaysActiveSheet(false);

            var j = TitleRow01 + 1;//itera según cada estándar
            var i = TitleRow02 + 1;//itera la celda para cada tarea

            while (!string.IsNullOrEmpty("" + stdCells.GetCell(3, j).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(stdCells.GetCell(1, j).Value2);
                    var workGroup = _cells.GetEmptyIfNull(stdCells.GetCell(2, j).Value2);
                    var stdJobNo = _cells.GetEmptyIfNull(stdCells.GetCell(3, j).Value2);

                    var taskList = StandardJobActions.FetchStandardJobTask(_eFunctions, districtCode, workGroup, stdJobNo);


                    foreach (var task in taskList)
                    {
                        //Para resetear el estilo
                        _cells.GetRange(1, i, ResultColumn02, i).Style = StyleConstants.Normal;
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + task.DistrictCode;
                        _cells.GetCell(2, i).Value = "" + task.WorkGroup;
                        _cells.GetCell(3, i).Value = "'" + task.StandardJob;
                        _cells.GetCell(4, i).Value = "" + task.StandardJobDescription;
                        //ACTION
                        _cells.GetCell(5, i).Value = "M";
                        //GENERAL
                        _cells.GetCell(6, i).Value = "'" + task.SjTaskNo;
                        _cells.GetCell(7, i).Value = "" + task.SjTaskDesc;
                        _cells.GetCell(8, i).Value = "'" + task.JobDescCode;
                        _cells.GetCell(9, i).Value = "'" + task.SafetyInstr;
                        _cells.GetCell(10, i).Value = "'" + task.CompleteInstr;
                        _cells.GetCell(11, i).Value = "'" + task.ComplTextCode;
                        //PLANNING

                        _cells.GetCell(12, i).Value = "" + task.AssignPerson;
                        _cells.GetCell(13, i).Value = "'" + task.EstimatedMachHrs;
                        _cells.GetCell(14, i).Value = "'" + task.UnitOfWork;
                        _cells.GetCell(15, i).Value = "" + task.UnitsRequired;
                        _cells.GetCell(16, i).Value = "" + task.UnitsPerDay;


                        //RECURSOS
                        _cells.GetCell(17, i).Value = "" + task.EstimatedDurationsHrs;
                        _cells.GetCell(18, i).Value = "" + task.NoLabor;
                        _cells.GetCell(19, i).Value = "" + task.NoMaterial;
                        //APL
                        _cells.GetCell(20, i).Value = "'" + task.AplEquipmentGrpId;
                        _cells.GetCell(21, i).Value = "'" + task.AplType;
                        _cells.GetCell(22, i).Value = "'" + task.AplCompCode;
                        _cells.GetCell(23, i).Value = "'" + task.AplCompModCode;
                        _cells.GetCell(24, i).Value = "'" + task.AplSeqNo;
                        _cells.GetRange(20, i, 24, i).Style = !string.IsNullOrWhiteSpace(task.AplType)
                            ? StyleConstants.Error : StyleConstants.Normal;

                        var stdTextId = "JI" + task.DistrictCode + task.StandardJob + task.SjTaskNo;
                        _cells.GetCell(25, i).Value = StdText.GetText(urlService, stOpContext, stdTextId);
                        _cells.GetCell(25, i).WrapText = false;
                        _cells.GetCell(3, i).Select();
                        i++;//aumenta tarea
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(1, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(1, j).Value2);
                    _cells.GetCell(2, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(2, j).Value2);
                    _cells.GetCell(3, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(3, j).Value2);
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn02, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:ReviewStandardJobTasks()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta std
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            _cells.SetCursorDefault();
        }

        private void QualityReviewStandardJobs()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRange(TableNameQualRev);
            var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value2);
            var workGroup = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value2);

            var i = TitleRow01 + 1;

            var listJobs = StandardJobActions.FetchStandardJob(_eFunctions, districtCode, workGroup, true);


            foreach (var stdJob in listJobs)
            {
                try
                {
                    //Para resetear el estilo
                    _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                    //GENERAL
                    _cells.GetCell(1, i).Value = "" + stdJob.DistrictCode;
                    _cells.GetCell(2, i).Value = "" + stdJob.WorkGroup;
                    _cells.GetCell(3, i).Value = "" + stdJob.StandardJobNo;
                    _cells.GetCell(4, i).Value = "" + stdJob.Status;
                    _cells.GetCell(5, i).Value = "" + stdJob.StandardJobDescription;
                    //CONSULTA                   
                    _cells.GetCell(6, i).Value = "" + stdJob.OrigPriority;
                    _cells.GetCell(6, i).Style = !WoTypeMtType.ValidatePriority(stdJob.OrigPriority)
                        ? StyleConstants.Error : StyleConstants.Normal;
                    _cells.GetCell(7, i).Value = "" + stdJob.WorkOrderType;
                    _cells.GetCell(8, i).Value = "" + stdJob.MaintenanceType;
                    _cells.GetRange(7, i, 8, i).Style = !WoTypeMtType.ValidateWoMtTypeCode(stdJob.WorkOrderType, stdJob.MaintenanceType)
                        ? StyleConstants.Error : StyleConstants.Normal;
                    _cells.GetCell(9, i).Value = "" + stdJob.UnitOfWork;
                    _cells.GetCell(10, i).Value = "" + stdJob.UnitsRequired;
                    if (!string.IsNullOrWhiteSpace(stdJob.UnitOfWork))
                        _cells.GetRange(9, i, 10, i).Style = int.Parse(stdJob.UnitsRequired) > 0
                            ? StyleConstants.Warning : StyleConstants.Error;
                    else
                        _cells.GetRange(9, i, 10, i).Style = StyleConstants.Normal;
                    _cells.GetCell(11, i).Value = "" + stdJob.CalculatedDurationsHrsFlg;
                    _cells.GetCell(12, i).Value = "" + stdJob.ResUpdateFlag;
                    _cells.GetCell(13, i).Value = "" + stdJob.MatUpdateFlag;
                    _cells.GetCell(14, i).Value = "" + stdJob.EquipmentUpdateFlag;
                    _cells.GetCell(15, i).Value = "" + stdJob.JobCode1;
                    _cells.GetCell(16, i).Value = "" + stdJob.JobCode2;
                    _cells.GetCell(17, i).Value = "" + stdJob.JobCode3;
                    _cells.GetCell(18, i).Value = "" + stdJob.JobCode4;
                    _cells.GetCell(19, i).Value = "" + stdJob.JobCode5;
                    _cells.GetCell(20, i).Value = "" + stdJob.JobCode6;
                    _cells.GetCell(21, i).Value = "" + stdJob.JobCode7;
                    _cells.GetCell(22, i).Value = "" + stdJob.JobCode8;
                    _cells.GetCell(23, i).Value = "" + stdJob.JobCode9;
                    _cells.GetCell(24, i).Value = "" + stdJob.JobCode10;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:QualityReviewStandardJobs()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(3, i).Select();
                    i++;
                }
            }

            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void CreateStandardJobList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableNameQualRev, ResultColumnQualRev);
            var i = TitleRowQualRev + 1;
            const int validationRow = TitleRow01 - 1;

            var opSheet = new StandardJobService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var stdJob = new StandardJob();
                    //GENERAL

                    stdJob.DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    stdJob.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    stdJob.StandardJobNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    stdJob.Status = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);
                    stdJob.StandardJobDescription = MyUtilities.IsTrue(_cells.GetCell(5, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value) : null;
                    //USO_OTS	        
                    //USO_MSTS	        
                    //ULTIMO_USO	    
                    //NO_OF_TASKS	    
                    stdJob.OriginatorId = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null;
                    stdJob.AssignPerson = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    stdJob.OrigPriority = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;
                    stdJob.WorkOrderType = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    stdJob.MaintenanceType = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null;
                    stdJob.CompCode = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    stdJob.CompModCode = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    stdJob.UnitOfWork = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null;
                    stdJob.UnitsRequired = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;
                    stdJob.CalculatedDurationsHrsFlg = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? "" + MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(19, i).Value)) : null;
                    stdJob.EstimatedDurationsHrs = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null;
                    stdJob.AccountCode = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null;
                    stdJob.ReallocAccCode = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    stdJob.ProjectNo = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null;
                    stdJob.EstimatedOtherCost = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    //CALC_LAB_HRS	    
                    //CALC_LAB_COST	    
                    //CALC_MAT_COST	    
                    //CALC_EQUIP_COST	
                    stdJob.JobCode1 = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;
                    stdJob.JobCode2 = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null;
                    stdJob.JobCode3 = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null;
                    stdJob.JobCode4 = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null;
                    stdJob.JobCode5 = MyUtilities.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null;
                    stdJob.JobCode6 = MyUtilities.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null;
                    stdJob.JobCode7 = MyUtilities.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null;
                    stdJob.JobCode8 = MyUtilities.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null;
                    stdJob.JobCode9 = MyUtilities.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null;
                    stdJob.JobCode10 = MyUtilities.IsTrue(_cells.GetCell(38, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value) : null;
                    stdJob.ExtText = MyUtilities.IsTrue(_cells.GetCell(39, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value) : null;
                    StandardJobActions.CreateStandardJob(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, stdJob);
                    if(!string.IsNullOrWhiteSpace(stdJob.ExtText))
                        StandardJobActions.SetStandardJobText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, stdJob);

                    _cells.GetCell(ResultColumn01, i).Value = "CREADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateStandardJobList()", ex.Message);
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
        private void UpdateStandardJobList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            var opSheet = new StandardJobService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var stdJob = new StandardJob();
                    //GENERAL

                    stdJob.DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    stdJob.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    stdJob.StandardJobNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    stdJob.Status = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);
                    stdJob.StandardJobDescription = MyUtilities.IsTrue(_cells.GetCell(5, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value) : null;
                    //USO_OTS	        
                    //USO_MSTS	        
                    //ULTIMO_USO	    
                    //NO_OF_TASKS	    
                    stdJob.OriginatorId = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null;
                    stdJob.AssignPerson = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    stdJob.OrigPriority = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;
                    stdJob.WorkOrderType = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    stdJob.MaintenanceType = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null;
                    stdJob.CompCode = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    stdJob.CompModCode = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    stdJob.UnitOfWork = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null;
                    stdJob.UnitsRequired = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;
                    stdJob.CalculatedDurationsHrsFlg = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? "" + MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(19, i).Value)) : null;
                    stdJob.EstimatedDurationsHrs = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null;
                    stdJob.AccountCode = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null;
                    stdJob.ReallocAccCode = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    stdJob.ProjectNo = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null;
                    stdJob.EstimatedOtherCost = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    //CALC_LAB_HRS	    
                    //CALC_LAB_COST	    
                    //CALC_MAT_COST	    
                    //CALC_EQUIP_COST	
                    stdJob.JobCode1 = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;
                    stdJob.JobCode2 = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null;
                    stdJob.JobCode3 = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null;
                    stdJob.JobCode4 = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null;
                    stdJob.JobCode5 = MyUtilities.IsTrue(_cells.GetCell(33, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null;
                    stdJob.JobCode6 = MyUtilities.IsTrue(_cells.GetCell(34, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null;
                    stdJob.JobCode7 = MyUtilities.IsTrue(_cells.GetCell(35, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(35, i).Value) : null;
                    stdJob.JobCode8 = MyUtilities.IsTrue(_cells.GetCell(36, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(36, i).Value) : null;
                    stdJob.JobCode9 = MyUtilities.IsTrue(_cells.GetCell(37, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(37, i).Value) : null;
                    stdJob.JobCode10 = MyUtilities.IsTrue(_cells.GetCell(38, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(38, i).Value) : null;

                    //Texto Extendido
                    stdJob.ExtText = MyUtilities.IsTrue(_cells.GetCell(39, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(39, i).Value) : null;

                    //
                    //stdJob.CalculatedDurationsHrsFlg = "true";
                    //stdJob.resUpdateFlag = "true";
                    //stdJob.matUpdateFlag = "true";
                    //stdJob.equipmentUpdateFlag = "true";
                    //stdJob.otherUpdateFlag = "true";
                    //

                    StandardJobActions.ModifyStandardJob(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, stdJob, true);
                    StandardJobActions.SetStandardJobText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, stdJob);

                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateStandardJobList()", ex.Message);
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
        private void UpdateQualityStandardList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableNameQualRev, ResultColumnQualRev);
            var i = TitleRowQualRev + 1;
            const int validationRow = TitleRowQualRev - 1;

            var opSheet = new StandardJobService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var stdJob = new StandardJob
                    {
                        DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        StandardJobNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        Status = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value),
                        OrigPriority = MyUtilities.IsTrue(_cells.GetCell(6, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value) : null,
                        WorkOrderType = MyUtilities.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null,
                        MaintenanceType = MyUtilities.IsTrue(_cells.GetCell(8, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value) : null,
                        UnitOfWork = MyUtilities.IsTrue(_cells.GetCell(9, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value) : null,
                        UnitsRequired = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null,
                        CalculatedDurationsHrsFlg = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? "" + MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(11, i).Value)) : null,
                        ResUpdateFlag = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? "" + MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(12, i).Value)) : null,
                        MatUpdateFlag = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? "" + MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(13, i).Value)) : null,
                        EquipmentUpdateFlag = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? "" + MyUtilities.IsTrue(_cells.GetEmptyIfNull(_cells.GetCell(14, i).Value)) : null,
                        JobCode1 = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null,
                        JobCode2 = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null,
                        JobCode3 = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null,
                        JobCode4 = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null,
                        JobCode5 = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null,
                        JobCode6 = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null,
                        JobCode7 = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null,
                        JobCode8 = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null,
                        JobCode9 = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null,
                        JobCode10 = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null,
                        ExtText = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value)
                    };


                    StandardJobActions.ModifyStandardJob(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, stdJob, true);

                    _cells.GetCell(ResultColumnQualRev, i).Value = "ACTUALIZADA";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumnQualRev, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnQualRev, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumnQualRev, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateQualityStandardList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumnQualRev, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        private void UpdateStandardJobStatus()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;

            var opSheet = new StandardJobService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var stdJob = new StandardJob
                    {
                        DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        StandardJobNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                    };

                    var resultStatus = StandardJobActions.UpdateStandardJobStatus(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), opSheet, stdJob, _standardStatus);

                    _cells.GetCell(ResultColumn01, i).Value = "ESTADO " + resultStatus;
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateStandardJobStatus()", ex.Message);
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

        private void ExecuteTaskActionsPost()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);
            var i = TitleRow02 + 1;

            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label, ServiceType.PostService);
            _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlService);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                    var stdTask = new StandardJobTask
                    {
                        DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        StandardJob = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        SjTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),
                        SjTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        JobDescCode = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        SafetyInstr = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value),
                        CompleteInstr = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                        ComplTextCode = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value),
                        AssignPerson = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value),
                        EstimatedMachHrs = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value),
                        EstimatedMachHrsSpecified = "Y",
                        UnitOfWork = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value),
                        UnitsRequired = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value),
                        UnitsRequiredSpecified = "Y",
                        UnitsPerDay = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value),
                        UnitsPerDaySpecified = "Y",
                        EstimatedDurationsHrs = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value),
                        EstimatedDurationsHrsSpecified = "Y",
                        NoLabor = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value),
                        NoMaterial = _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value),
                        AplEquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value),
                        AplType = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value),
                        AplCompCode = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value),
                        AplCompModCode = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value),
                        AplSeqNo = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value),
                        ExtTaskText = _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value)
                    };

                    if (string.IsNullOrWhiteSpace(action))
                        continue;

                    if (action.Equals("M"))
                    {
                        //StandardJobActions.ModifyStandardJobTaskPost(_eFunctions, stdTask, true);
                        //var opC = new StandardJobTaskService.OperationContext()
                        //{
                        //    district = _frmAuth.EllipseDsct,
                        //    position = _frmAuth.EllipsePost,
                        //    maxInstances = 100,
                        //    maxInstancesSpecified = true,
                        //    returnWarnings = Debugger.DebugWarnings,
                        //    returnWarningsSpecified = true
                        //};
                        StandardJobActions.ModifyStandardJobTaskPost(_eFunctions, stdTask);
                        StandardJobActions.SetStandardJobTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, stdTask);
                    }
                    else if (action.Equals("C"))
                    {
                        StandardJobActions.CreateStandardJobTaskPost(_eFunctions, stdTask);
                        StandardJobActions.SetStandardJobTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, stdTask);
                    }
                    else
                        continue;

                    _cells.GetCell(ResultColumn02, i).Value = "OK";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteTaskActionsPost()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn02, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        /// <summary>
        /// Ejecuta las acciones de tarea mediante el servicio EWS
        /// </summary>
        // ReSharper disable once UnusedMember.Local
        private void ExecuteTaskActions()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);
            var i = TitleRow02 + 1;

            var opSheet = new StandardJobTaskService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);


            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                    var stdTask = new StandardJobTask
                    {
                        DistrictCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value),
                        WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value),
                        StandardJob = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value),
                        SjTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value),
                        SjTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                        JobDescCode = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value),
                        SafetyInstr = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value),
                        CompleteInstr = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value),
                        ComplTextCode = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value),
                        AssignPerson = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value),
                        EstimatedMachHrs = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value),
                        UnitOfWork = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value),
                        UnitsRequired = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value),
                        UnitsPerDay = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value),
                        EstimatedDurationsHrs = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value),
                        NoLabor = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value),
                        NoMaterial = _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value),
                        AplEquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value),
                        AplType = _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value),
                        AplCompCode = _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value),
                        AplCompModCode = _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value),
                        AplSeqNo = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value),
                        ExtTaskText = _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value)
                    };

                    if (string.IsNullOrWhiteSpace(action))
                        continue;

                    if (action.Equals("M"))
                    {
                        StandardJobActions.ModifyStandardJobTask(urlService, opSheet, stdTask, true);
                        StandardJobActions.SetStandardJobTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, stdTask);
                    }
                    else if (action.Equals("C"))
                    {
                        StandardJobActions.CreateStandardJobTask(urlService, opSheet, stdTask, true);
                        StandardJobActions.SetStandardJobTaskText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), _frmAuth.EllipseDsct, _frmAuth.EllipsePost, true, stdTask);
                    }
                    else
                        continue;

                    _cells.GetCell(ResultColumn02, i).Value = "OK";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteTaskActions()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn02, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void ReviewRequirements()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRange(TableName03);

            var taskCells = new ExcelStyleCells(_excelApp, SheetName02);
            taskCells.SetAlwaysActiveSheet(false);

            var j = TitleRow02 + 1;//itera según cada tarea
            var i = TitleRow03 + 1;//itera la celda para cada requerimiento

            while (!string.IsNullOrEmpty("" + taskCells.GetCell(3, j).Value) && !string.IsNullOrEmpty("" + taskCells.GetCell(6, j).Value))
            {
                try
                {
                    var districtCode = _cells.GetEmptyIfNull(taskCells.GetCell(1, j).Value2);
                    var workGroup = _cells.GetEmptyIfNull(taskCells.GetCell(2, j).Value2);
                    var stdJobNo = _cells.GetEmptyIfNull(taskCells.GetCell(3, j).Value2);
                    var taskNo = _cells.GetEmptyIfNull(taskCells.GetCell(6, j).Value2);

                    var reqList = StandardJobActions.FetchTaskRequirements(_eFunctions, districtCode, workGroup, stdJobNo, taskNo);

                    foreach (var req in reqList)
                    {
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + req.districtCode;
                        _cells.GetCell(2, i).Value = "" + req.workGroup;
                        _cells.GetCell(3, i).Value = "" + req.standardJob;
                        _cells.GetCell(4, i).Value = "" + req.sJTaskNo;
                        _cells.GetCell(5, i).Value = "" + req.sJTaskDesc;
                        _cells.GetCell(6, i).Value = "M";
                        _cells.GetCell(7, i).Value = "" + req.reqType;
                        _cells.GetCell(8, i).Value = "" + req.seqNo;
                        _cells.GetCell(9, i).Value = "" + req.reqCode;
                        _cells.GetCell(10, i).Value = "" + req.reqDesc;
                        _cells.GetCell(11, i).Value = "" + req.qtyReq;
                        _cells.GetCell(12, i).Value = "" + req.hrsReq;
                        _cells.GetCell(ResultColumn03, i).Select();
                        i++;//aumenta req
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(1, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(1, j).Value2);
                    _cells.GetCell(2, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(2, j).Value2);
                    _cells.GetCell(3, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(3, j).Value2);
                    _cells.GetCell(4, i).Value = _cells.GetEmptyIfNull(_cells.GetCell(6, j).Value2);
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn03, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:ReviewRequirements()", ex.Message);
                    i++;
                }
                finally
                {
                    j++;//aumenta task
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        private void ExecuteRequirementActions()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName03, ResultColumn03);
            var i = TitleRow03 + 1;

            var opSheetResource = new ResourceReqmntsService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };
            var opSheetMaterial = new MaterialReqmntsService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };
            var opSheetEquipment = new EquipmentReqmntsService.OperationContext()
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true,
                maxInstancesSpecified = true
            };


            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value) && !string.IsNullOrEmpty("" + _cells.GetCell(4, i).Value))
            {
                try
                {
                    // ReSharper disable once UseObjectOrCollectionInitializer
                    var taskReq = new TaskRequirement();
                    //GENERAL

                    taskReq.districtCode = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    taskReq.workGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    taskReq.standardJob = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                    //STD_JOB_DESC	
                    taskReq.sJTaskNo = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value);
                    taskReq.sJTaskNo = string.IsNullOrWhiteSpace(taskReq.sJTaskNo) ? "001" : taskReq.sJTaskNo;
                    taskReq.sJTaskDesc = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                    string action = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value);
                    taskReq.reqType = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value);
                    taskReq.seqNo = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value);
                    taskReq.reqCode = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value);
                    taskReq.reqDesc = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value);
                    taskReq.qtyReq = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value);
                    taskReq.hrsReq = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value);
                    taskReq.uoM = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value);

                    if (string.IsNullOrWhiteSpace(action))
                        continue;
                    else if (action.Equals("C"))
                    {
                        if (taskReq.reqType.Equals("LAB"))
                            StandardJobActions.CreateTaskResource(urlService, opSheetResource, taskReq);
                        else if (taskReq.reqType.Equals("MAT"))
                            StandardJobActions.CreateTaskMaterial(urlService, opSheetMaterial, taskReq);
                        else if (taskReq.reqType.Equals("EQU"))
                            StandardJobActions.CreateTaskEquipment(urlService, opSheetEquipment, taskReq);
                    }
                    else if (action.Equals("M"))
                    {
                        if (taskReq.reqType.Equals("LAB"))
                            StandardJobActions.ModifyTaskResource(urlService, opSheetResource, taskReq);
                        else if (taskReq.reqType.Equals("MAT"))
                            StandardJobActions.ModifyTaskMaterial(urlService, opSheetMaterial, taskReq);
                        else if (taskReq.reqType.Equals("EQU"))
                            StandardJobActions.ModifyTaskEquipment(urlService, opSheetEquipment, taskReq);
                    }
                    else if (action.Equals("D"))
                    {
                        if (taskReq.reqType.Equals("LAB"))
                            StandardJobActions.DeleteTaskResource(urlService, opSheetResource, taskReq);
                        else if (taskReq.reqType.Equals("MAT"))
                            StandardJobActions.DeleteTaskMaterial(urlService, opSheetMaterial, taskReq);
                        else if (taskReq.reqType.Equals("EQU"))
                            StandardJobActions.DeleteTaskEquipment(urlService, opSheetEquipment, taskReq);
                    }


                    _cells.GetCell(ResultColumn03, i).Value = "OK";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ExecuteRequirementActions()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn03, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        public void GetAplTaskRequirements()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var taskCells = new ExcelStyleCells(_excelApp, SheetName02);
            taskCells.SetAlwaysActiveSheet(false);

            var j = TitleRow02 + 1;//itera según cada tarea
            var i = TitleRow03 + 1;//itera la celda para cada requerimiento

            while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value) && !string.IsNullOrEmpty("" + _cells.GetCell(4, i).Value))
                i++;

            while (!string.IsNullOrEmpty("" + taskCells.GetCell(3, j).Value) && !string.IsNullOrEmpty("" + taskCells.GetCell(6, j).Value))
            {
                var districtCode = _cells.GetEmptyIfNull(taskCells.GetCell(1, j).Value2);
                var workGroup = _cells.GetEmptyIfNull(taskCells.GetCell(2, j).Value2);
                var stdJobNo = _cells.GetEmptyIfNull(taskCells.GetCell(3, j).Value2);
                var taskNo = _cells.GetEmptyIfNull(taskCells.GetCell(6, j).Value2);
                var taskDesc = _cells.GetEmptyIfNull(taskCells.GetCell(7, j).Value2);
                var aplEgi = _cells.GetEmptyIfNull(taskCells.GetCell(20, j).Value);
                var aplType = _cells.GetEmptyIfNull(taskCells.GetCell(21, j).Value);
                var aplCompCode = _cells.GetEmptyIfNull(taskCells.GetCell(22, j).Value);
                var aplCompModCode = _cells.GetEmptyIfNull(taskCells.GetCell(23, j).Value);
                var seqNo = _cells.GetEmptyIfNull(taskCells.GetCell(24, j).Value);

                if (string.IsNullOrWhiteSpace(aplEgi) && string.IsNullOrWhiteSpace(aplType))
                {
                    j++;
                    continue;
                }

                try
                {
                    var sqlQuery = Queries.GetAplRequirementsQuery(_eFunctions.dbReference, _eFunctions.dbLink, aplEgi, aplType, aplCompCode, aplCompModCode, seqNo);

                    var reqDataReader = _eFunctions.GetQueryResult(sqlQuery);

                    if (reqDataReader == null || reqDataReader.IsClosed || !reqDataReader.HasRows)
                        continue;

                    while (reqDataReader.Read())
                    {
                        //GENERAL
                        _cells.GetCell(1, i).Value = "" + districtCode;
                        _cells.GetCell(2, i).Value = "" + workGroup;
                        _cells.GetCell(3, i).Value = "" + stdJobNo;
                        _cells.GetCell(4, i).Value = "" + taskNo;
                        _cells.GetCell(5, i).Value = "" + taskDesc;
                        _cells.GetCell(6, i).Value = "C";
                        _cells.GetCell(7, i).Value = "MAT";
                        _cells.GetCell(8, i).Value = "";
                        _cells.GetCell(9, i).Value = "" + reqDataReader["STOCK_CODE"].ToString().Trim();
                        _cells.GetCell(10, i).Value = "" + reqDataReader["ITEM_DESC"].ToString().Trim();
                        _cells.GetCell(11, i).Value = "" + reqDataReader["QTY_REQUIRED"].ToString().Trim();
                        _cells.GetCell(12, i).Value = "" + reqDataReader["QTY_INSTALLED"].ToString().Trim();
                        _cells.GetCell(ResultColumn03, i).Select();
                        i++;//aumenta req
                    }
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(1, i).Value = districtCode;
                    _cells.GetCell(2, i).Value = workGroup;
                    _cells.GetCell(3, i).Value = stdJobNo;
                    _cells.GetCell(4, i).Value = taskNo;
                    _cells.GetCell(ResultColumn03, i).Value = "ERROR: " + ex.Message;
                    _cells.GetCell(ResultColumn03, i).Select();
                    Debugger.LogError("RibbonEllipse.cs:GetAplTaskRequirements()", ex.Message);
                    i++;
                }
                finally
                {
                    _eFunctions.CloseConnection();
                    j++;//aumenta task
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        public static class Queries
        {
            public static string GetAplRequirementsQuery(string dbReference, string dbLink, string aplEgi, string aplType, string aplCompCode, string aplCompModCode, string seqNo)
            {
                if (string.IsNullOrWhiteSpace(aplCompCode))
                    aplCompCode = " IS NULL";
                else
                    aplCompCode = " = '" + aplCompCode + "'";

                if (string.IsNullOrWhiteSpace(aplCompModCode))
                    aplCompModCode = " IS NULL";
                else
                    aplCompModCode = " = '" + aplCompModCode + "'";

                var sqlQuery = "" +
                    " SELECT" +
                    "   AST.EQUIP_GRP_ID , AST.APL_TYPE, AST.COMP_CODE, AST.COMP_MOD_CODE, AST.APL_SEQ_NO, AST.APL_ITEM_NUM, AST.PART_NO, AST.MNEMONIC, AST.STOCK_CODE, AST.ITEM_DESC, AST.QTY_REQUIRED, AST.QTY_INSTALLED" +
                    " FROM" +
                    "   " + dbReference + ".MSF131" + dbLink + "  AST" +
                    " WHERE" +
                    "   TRIM(AST.EQUIP_GRP_ID) = '" + aplEgi + "' AND AST.APL_SEQ_NO = '" + seqNo + "' AND AST.APL_TYPE = '" + aplType + "' AND TRIM(AST.COMP_CODE) " + aplCompCode + " AND TRIM(AST.COMP_MOD_CODE) " + aplCompModCode + "";

                return sqlQuery;
            }

            public static string FetchReferenceCodeItems(string dbReference, string dbLink, string entityType, string entityValue, string refNo, string seqNum = null)
            {
                if (!string.IsNullOrWhiteSpace(refNo))
                    refNo = " AND RC.REF_NO = '" + refNo.PadLeft(3, '0') + "'";
                if (!string.IsNullOrWhiteSpace(seqNum))
                    seqNum = " AND RC.SEQ_NUM = '" + seqNum.PadLeft(3, '0') + "'";
                var query = "" +
                            " SELECT RC.ENTITY_TYPE, " +
                            "   RC.ENTITY_VALUE, " +
                            "   RC.REF_NO, " +
                            "   RC.SEQ_NUM, " +
                            "   RC.REF_CODE, " +
                            "   RCE.FIELD_TYPE, " +
                            "   RCE.SHORT_NAMES, " +
                            "   RCE.SCREEN_LITERAL, " +
                            "   RC.STD_TXT_KEY, " +
                            "   RCE.STD_TEXT_FLAG " +
                            " FROM " +
                            "     " + dbReference + ".MSF071" + dbLink + " RC LEFT JOIN " + dbReference + ".MSF070" + dbLink + " RCE " +
                            "         ON (RC.ENTITY_TYPE = RCE.ENTITY_TYPE AND RC.REF_NO = RCE.REF_NO) " +
                            " WHERE RCE.ENTITY_TYPE = '" + entityType + "' " +
                            " AND RC.ENTITY_VALUE = '" + entityValue + "' " +
                            " " + refNo +
                            " " + seqNum;
                return query;
            }
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                if(_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }

        }

        private void btnReviewStandardReferenceCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName05))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                Debugger.LogError("EllipseStandardJobsExcelAddIn:RibbonEllipse.cs:ReviewStandardReferenceCodes()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void ReviewRefCodesList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            var opSheet = new StandardJobService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var i = TitleRow05 + 1;

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var district = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                    district = string.IsNullOrWhiteSpace(district) ? "ICOR" : district;
                    var workGroup = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value);
                    var standardJobNo = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);


                    _cells.GetCell(1, TitleRow05).Value = "DISTRICT";
                    _cells.GetCell(2, TitleRow05).Value = "WORK_GROUP";
                    _cells.GetCell(3, TitleRow05).Value = "STD_JOB_NO";

                    var std = StandardJobActions.FetchStandardJob(_eFunctions, district, workGroup, standardJobNo);
                    if (std.StandardJobNo == null)
                        throw new Exception("ESTANDAR NO ENCONTRADO");

                    StandardJobReferenceCodes stdRefCodes = StandardJobActions.GetStandardJobReferenceCodes(_eFunctions, urlService, opSheet, district, standardJobNo);
                    //GENERAL
                    _cells.GetCell(4, i).Value = "'" + std.StandardJobDescription;
                    _cells.GetCell(5, i).Value = "'" + stdRefCodes.WorkRequest;
                    _cells.GetCell(6, i).Value = "'" + stdRefCodes.ComentariosDuraciones;
                    _cells.GetCell(7, i).Value = "'" + stdRefCodes.ComentariosDuracionesText;
                    _cells.GetCell(8, i).Value = "'" + stdRefCodes.EmpleadoId;
                    _cells.GetCell(9, i).Value = "'" + stdRefCodes.NroComponente;
                    _cells.GetCell(10, i).Value = "'" + stdRefCodes.P1EqLivMed;
                    _cells.GetCell(11, i).Value = "'" + stdRefCodes.P2EqMovilMinero;
                    _cells.GetCell(12, i).Value = "'" + stdRefCodes.P3ManejoSustPeligrosa;
                    _cells.GetCell(13, i).Value = "'" + stdRefCodes.P4GuardasEquipo;
                    _cells.GetCell(14, i).Value = "'" + stdRefCodes.P5Aislamiento;
                    _cells.GetCell(15, i).Value = "'" + stdRefCodes.P6TrabajosAltura;
                    _cells.GetCell(16, i).Value = "'" + stdRefCodes.P7ManejoCargas;
                    _cells.GetCell(17, i).Value = "'" + stdRefCodes.ProyectoIcn;
                    _cells.GetCell(18, i).Value = "'" + stdRefCodes.Reembolsable;
                    _cells.GetCell(19, i).Value = "'" + stdRefCodes.FechaNoConforme;
                    _cells.GetCell(20, i).Value = "'" + stdRefCodes.FechaNoConformeText;
                    _cells.GetCell(21, i).Value = "'" + stdRefCodes.NoConforme;
                    _cells.GetCell(22, i).Value = "'" + stdRefCodes.FechaEjecucion;
                    _cells.GetCell(23, i).Value = "'" + stdRefCodes.HoraIngreso;
                    _cells.GetCell(24, i).Value = "'" + stdRefCodes.HoraSalida;
                    _cells.GetCell(25, i).Value = "'" + stdRefCodes.NombreBuque;
                    _cells.GetCell(26, i).Value = "'" + stdRefCodes.CalificacionEncuesta;
                    _cells.GetCell(27, i).Value = "'" + stdRefCodes.TareaCritica;
                    _cells.GetCell(28, i).Value = "'" + stdRefCodes.Garantia;
                    _cells.GetCell(29, i).Value = "'" + stdRefCodes.GarantiaText;
                    _cells.GetCell(30, i).Value = "'" + stdRefCodes.CodigoCertificacion;
                    _cells.GetCell(31, i).Value = "'" + stdRefCodes.FechaEntrega;
                    _cells.GetCell(32, i).Value = "'" + stdRefCodes.RelacionarEv;
                    _cells.GetCell(33, i).Value = "'" + stdRefCodes.Departamento;
                    _cells.GetCell(34, i).Value = "'" + stdRefCodes.Localizacion;

                    _cells.GetCell(ResultColumn05, i).Value = "CONSULTADO";
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("EllipseStandardJobsExcelAddIn:RibbonEllipse.cs:ReviewRefCodesList()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(2, i).Select();
                    i++;
                    _eFunctions.CloseConnection();
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        private void btnUpdateStandardReferenceCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName05))
                {
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(UpdateReferenceCodes);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:UpdateStandardReferenceCodes()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        public void UpdateReferenceCodes()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName05, ResultColumn05);

            var i = TitleRow05 + 1;
            const int validationRow = TitleRow05 - 1;

            var opSheet = new StandardJobService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
            {
                try
                {
                    //GENERAL
                    var district = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(1, i).Value);
                    var standardJob = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(3, i).Value);
                    var sjRefCodes = new StandardJobReferenceCodes()
                    {
                        WorkRequest = MyUtilities.IsTrue(_cells.GetCell(5, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value) : null,
                        ComentariosDuraciones = MyUtilities.IsTrue(_cells.GetCell(6, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value) : null,
                        ComentariosDuracionesText = MyUtilities.IsTrue(_cells.GetCell(7, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null,
                        EmpleadoId = MyUtilities.IsTrue(_cells.GetCell(8, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value) : null,
                        NroComponente = MyUtilities.IsTrue(_cells.GetCell(9, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value) : null,
                        P1EqLivMed = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null,
                        P2EqMovilMinero = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null,
                        P3ManejoSustPeligrosa = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null,
                        P4GuardasEquipo = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null,
                        P5Aislamiento = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null,
                        P6TrabajosAltura = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null,
                        P7ManejoCargas = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null,
                        ProyectoIcn = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null,
                        Reembolsable = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null,
                        FechaNoConforme = MyUtilities.IsTrue(_cells.GetCell(19, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(19, i).Value) : null,
                        FechaNoConformeText = MyUtilities.IsTrue(_cells.GetCell(20, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(20, i).Value) : null,
                        NoConforme = MyUtilities.IsTrue(_cells.GetCell(21, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(21, i).Value) : null,
                        FechaEjecucion = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null,
                        HoraIngreso = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null,
                        HoraSalida = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null,
                        NombreBuque = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null,
                        CalificacionEncuesta = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null,
                        TareaCritica = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null,
                        Garantia = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null,
                        GarantiaText = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null,
                        CodigoCertificacion = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null,
                        FechaEntrega = MyUtilities.IsTrue(_cells.GetCell(31, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(31, i).Value) : null,
                        RelacionarEv = MyUtilities.IsTrue(_cells.GetCell(32, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(32, i).Value) : null,
                        Departamento = MyUtilities.IsTrue(_cells.GetCell(33, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(33, i).Value) : null,
                        Localizacion = MyUtilities.IsTrue(_cells.GetCell(34, validationRow).Value)
                                ? _cells.GetEmptyIfNull(_cells.GetCell(34, i).Value) : null
                    };


                    StandardJobActions.UpdateWorkOrderReferenceCodes(_eFunctions, urlService, opSheet, district, standardJob, sjRefCodes);

                    _cells.GetCell(ResultColumn05, i).Value = "ACTUALIZADO";
                    _cells.GetCell(1, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn05, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateReferenceCodes()", ex.Message);
                }
                finally
                {
                    _cells.GetCell(ResultColumn05, i).Select();
                    i++;
                }
            }
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if(_cells != null) _cells.SetCursorDefault();
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }
    }
}
