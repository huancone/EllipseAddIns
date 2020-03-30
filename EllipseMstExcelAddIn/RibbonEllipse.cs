using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseMaintSchedTaskClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Utilities;
using EllipseCommonsClassLibrary.Constants;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Threading;
using System.Web.Services.Ellipse;

// ReSharper disable UseObjectOrCollectionInitializer

namespace EllipseMstExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]

    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Application _excelApp;

        private const string SheetName01 = "MaintSchedTask";
        private const string SheetName02 = "ReScheduleMSTs";
        private const int TitleRow01 = 9;
        private const int TitleRow02 = 9;
        private const int ResultColumn01 = 30;
        private const int ResultColumn02 = 16;
        private const string TableName01 = "MstTable";
        private const string TableName02 = "ReScheduleTable";
        private const string ValidationSheetName = "ValidationSheetMst";

        private Thread _thread;


        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;

            var environments = Environments.GetEnvironmentList();
            foreach (var env in environments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnvironment.Items.Add(item);
            }

        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();
        }

        private void btnReviewMsts_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReviewMstList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnReReviewMst_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ReReviewMstList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnCreateMsts_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(CreateMstList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnDeleteMsts_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                var dr =
                    MessageBox.Show(@"Esta acción eliminará las tareas existentes. ¿Está seguro que desea continuar?",
                        @"ELIMINAR MST", MessageBoxButtons.YesNo);
                if (dr != DialogResult.Yes) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(DeleteMstList);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnUpdateMst_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => UpdateMstList(false));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnUpdateMstPost_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(() => UpdateMstList(true));

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnModifyNextSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ModifyNextScheduleList);

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

                _cells.GetCell("B1").Value = "MAINTENANCE SCHEDULE TASK - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "GRUPO";
                _cells.GetCell("A5").Value = "INDICADOR";
                _cells.GetCell("B5").Value = " A - " + MstIndicatorList.Active;

                _cells.GetRange("A3", "A5").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("B3", "B5").Style = _cells.GetStyle(StyleConstants.Select);

                //validaciones de encabezado
                _cells.SetValidationList(_cells.GetCell("B3"), Districts.GetDistrictList(), ValidationSheetName,
                    1);
                _cells.SetValidationList(_cells.GetCell("B4"),
                    Groups.GetWorkGroupList().Select(g => g.Name).ToList(), ValidationSheetName, 2, false);
                var listIndicator = MstIndicatorList.GetIndicatorsList();
                listIndicator.Add(" A - " + MstIndicatorList.Active);
                _cells.SetValidationList(_cells.GetCell("B5"), listIndicator, ValidationSheetName, 3);


                _cells.GetRange(1, TitleRow01, ResultColumn01 - 1, TitleRow01).Style =
                    _cells.GetStyle(StyleConstants.TitleRequired);
                for (var i = 7; i < ResultColumn01; i++)
                {
                    _cells.GetCell(i, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                    _cells.GetCell(i, TitleRow01 - 1).AddComment("Solo se modificará este campo si es verdadero (VERDADERO, TRUE, Y, 1)");
                    _cells.GetCell(i, TitleRow01 - 1).Value = "true";
                }
                //MST INFORMATION
                _cells.GetRange(1, TitleRow01 - 2, 6, TitleRow01 - 1).Style = StyleConstants.Select;
                _cells.GetCell(1, TitleRow01 - 2).Value = "MST";
                _cells.GetRange(1, TitleRow01 - 2, 6, TitleRow01 - 1).Merge();

                _cells.GetCell(1, TitleRow01).Value = "GRUPO DE TRABAJO";
                _cells.GetCell(2, TitleRow01).Value = "TIPO";
                _cells.GetCell(2, TitleRow01).AddComment("ES Equipos \nGS EGI");
                _cells.GetCell(3, TitleRow01).Value = "EQUIPO/EGI";
                _cells.GetCell(4, TitleRow01).Value = "COD.COMP.";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, TitleRow01).Value = "MOD.COMP.";
                _cells.GetCell(5, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(6, TitleRow01).Value = "NRO. MST";


                //GENERAL INFORMATION
                _cells.GetRange(7, TitleRow01 - 2, 11, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetCell(7, TitleRow01 - 2).Value = "GENERAL";
                _cells.GetRange(7, TitleRow01 - 2, 11, TitleRow01 - 2).Merge();

                _cells.GetCell(7, TitleRow01).Value = "DESCRIPCIÓN 1";
                _cells.GetCell(8, TitleRow01).Value = "DESCRIPCIÓN 2";
                _cells.GetCell(9, TitleRow01).Value = "ESTÁNDAR JOB";
                _cells.GetCell(10, TitleRow01).Value = "JOB DESC CODE";
                _cells.GetCell(11, TitleRow01).Value = "ASIGNADA A";
                //SCHEDULE INFORMATION
                _cells.GetRange(12, TitleRow01 - 2, 21, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetCell(12, TitleRow01 - 2).Value = "SCHEDULE";
                _cells.GetRange(12, TitleRow01 - 2, 21, TitleRow01 - 2).Merge();

                _cells.GetCell(12, TitleRow01).Value = "INDICADOR";
                _cells.GetCell(13, TitleRow01).Value = "FRECUENCIA";
                _cells.GetCell(14, TitleRow01).Value = "TIPO ESTAD.";
                _cells.GetCell(14, TitleRow01).AddComment("Ej: HR");
                _cells.GetCell(15, TitleRow01).Value = "ÚLT. ESTAD. PROG./LSS";
                _cells.GetCell(16, TitleRow01).Value = "ÚLT. ESTAD. RLZD./LPS";
                _cells.GetCell(17, TitleRow01).Value = "ÚLT. FECHA PROG./LSD";
                _cells.GetCell(17, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(18, TitleRow01).Value = "ÚLT. FECHA RLZD./LPD";
                _cells.GetCell(18, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetRange(15, TitleRow01, 18, TitleRow01).Style = StyleConstants.TitleOptional;

                _cells.GetCell(19, TitleRow01).Value = "SIG. FECHA PROG.";
                _cells.GetCell(19, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(19, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;
                _cells.GetCell(20, TitleRow01).Value = "SIG. ESTAD PROG.";
                _cells.GetCell(20, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(20, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;

                _cells.GetCell(21, TitleRow01).Value = "DESC. FRECUENCIA";
                _cells.GetCell(21, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(21, TitleRow01 - 1).Style = StyleConstants.ItalicSmall;


                _cells.GetRange(22, TitleRow01 - 2, 24, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetCell(22, TitleRow01 - 2).Value = "NEXT SCHEDULE";
                _cells.GetCell(22, TitleRow01 - 2).AddComment("Para indicadores de 1 - 4:\n 1. Last Sched Date\n 2. Last Sched Sat\n 3. Last Perf Date\n 4. Last Perf Stat");
                _cells.GetRange(22, TitleRow01 - 2, 24, TitleRow01 - 2).Merge();
                _cells.GetCell(22, TitleRow01).Value = "STAT TYPE";
                _cells.GetCell(23, TitleRow01).Value = "STAT VALUE";
                _cells.GetCell(24, TitleRow01).Value = "NEXT SCHED DATE";

                _cells.GetRange(25, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2).Style = StyleConstants.Select;
                _cells.GetCell(25, TitleRow01 - 2).Value = "FIXED SCHEDULE";
                _cells.GetCell(25, TitleRow01 - 2).AddComment("Para indicadores 7 y 8:\n 7. Fixed Date\n 8. Fixed Day");
                _cells.GetRange(25, TitleRow01 - 2, ResultColumn01 - 1, TitleRow01 - 2).Merge();
                _cells.GetCell(25, TitleRow01).Value = "OCCURRENCE TYPE";
                _cells.GetCell(25, TitleRow01).AddComment("Indicador 8. Indica cuál semana del mes\n1.Primera\n2.Segunda\n3.Tercera\n4.Cuarta\nL.Última");
                _cells.GetCell(26, TitleRow01).Value = "DAY OF WEEK";
                _cells.GetCell(26, TitleRow01).AddComment("Indicador 8. Indica cuál día de la semana elegida en OCCURRENCE TYPE\n1. Lunes\n2.Martes\n3.Miércoles\n4.Jueves\n5.Viernes\n6.Sábado\n7.Domingo");
                _cells.GetCell(27, TitleRow01).Value = "DAY OF MONTH";
                _cells.GetCell(27, TitleRow01).AddComment("Indicador 7. Indica cuál día del mes (número)");
                _cells.GetCell(28, TitleRow01).Value = "START MONTH";
                _cells.GetCell(28, TitleRow01).AddComment("Indicador 7 y 8. Indica a partir de qué mes\n1.Enero\n2.Febrero\n3.Marzo\n4.Abril\n5.Mayo\n6.Junio\n7.Julio\n8.Agosto\n24.Septiembre\n10.Octubre\n11.Noviembre\n12.Diciembre");
                _cells.GetCell(29, TitleRow01).Value = "START YEAR";

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);


                //asigno la validación de celda
                //

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat =
                    NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //HOJA 2 - RESCHEDULE
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("C1").Value = "RESCHEDULE MAINTENANCE SCHEDULE TASK - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetCell("B3").Style = _cells.GetStyle(StyleConstants.Select);

                //MST INFORMATION
                _cells.GetRange(1, TitleRow02 - 1, 5, TitleRow02 - 1).Style = StyleConstants.Select;
                _cells.GetCell(1, TitleRow02 - 1).Value = "MST";
                _cells.GetRange(1, TitleRow02 - 1, 5, TitleRow02 - 1).Merge();

                _cells.GetCell(1, TitleRow02).Value = "WORKGROUP";
                _cells.GetCell(2, TitleRow02).Value = "EQUIPMENT";
                _cells.GetCell(2, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(3, TitleRow02).Value = "COD.COMP.";
                _cells.GetCell(3, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(4, TitleRow02).Value = "MOD.COMP.";
                _cells.GetCell(4, TitleRow02).Style = StyleConstants.TitleOptional;
                _cells.GetCell(5, TitleRow02).Value = "MST NO.";
                _cells.GetCell(5, TitleRow02).Style = StyleConstants.TitleRequired;

                _cells.GetCell(6, TitleRow02 - 1).Value = "INDICATOR";
                _cells.GetCell(6, TitleRow02 - 1).Style = StyleConstants.Select;
                _cells.GetCell(6, TitleRow02).Value = "SCHED IND";
                _cells.GetCell(6, TitleRow02).AddComment("1. Last Sched Date\n 2. Last Sched Sat\n 3. Last Perf Date\n 4. Last Perf Stat\n 7. Fixed Date\n 8. Fixed Day");
                _cells.GetCell(6, TitleRow02).Style = StyleConstants.TitleRequired;

                _cells.GetRange(7, TitleRow02 - 1, 9, TitleRow02 - 1).Style = StyleConstants.Select;
                _cells.GetCell(7, TitleRow02 - 1).Value = "NEXT SCHEDULE";
                _cells.GetCell(7, TitleRow02 - 1).AddComment("Para indicadores de 1 - 4:\n 1. Last Sched Date\n 2. Last Sched Sat\n 3. Last Perf Date\n 4. Last Perf Stat");
                _cells.GetRange(7, TitleRow02 - 1, 9, TitleRow02 - 1).Merge();
                _cells.GetCell(7, TitleRow02).Value = "STAT TYPE";
                _cells.GetCell(8, TitleRow02).Value = "STAT VALUE";
                _cells.GetCell(9, TitleRow02).Value = "NEXT SCHED DATE";
                _cells.GetCell(9, TitleRow02).AddComment("YYYYMMDD");

                _cells.GetRange(10, TitleRow02 - 1, ResultColumn02 - 1, TitleRow02 - 1).Style = StyleConstants.Select;
                _cells.GetCell(10, TitleRow02 - 1).Value = "FIXED SCHEDULE";
                _cells.GetCell(10, TitleRow02 - 1).AddComment("Para indicadores 7 y 8:\n 7. Fixed Date\n 8. Fixed Day");
                _cells.GetRange(10, TitleRow02 - 1, ResultColumn02 - 1, TitleRow02 - 1).Merge();
                _cells.GetCell(10, TitleRow02).Value = "OCCURRENCE TYPE";
                _cells.GetCell(10, TitleRow02).AddComment("Indicador 8. Indica cuál semana del mes\n1.Primera\n2.Segunda\n3.Tercera\n4.Cuarta\nL.Última");
                _cells.GetCell(11, TitleRow02).Value = "DAY OF WEEK";
                _cells.GetCell(11, TitleRow02).AddComment("Indicador 8. Indica cuál día de la semana elegida en OCCURRENCE TYPE\n1. Lunes\n2.Martes\n3.Miércoles\n4.Jueves\n5.Viernes\n6.Sábado\n7.Domingo");
                _cells.GetCell(12, TitleRow02).Value = "DAY OF MONTH";
                _cells.GetCell(12, TitleRow02).AddComment("Indicador 7. Indica cuál día del mes (número)");
                _cells.GetCell(13, TitleRow02).Value = "FREQUENCY";
                _cells.GetCell(13, TitleRow02).AddComment("Indicador 7 y 8. Indica la frecuencia del mes (cada cuántos meses)");
                _cells.GetCell(14, TitleRow02).Value = "START MONTH";
                _cells.GetCell(14, TitleRow02).AddComment("Indicador 7 y 8. Indica a partir de qué mes\n1.Enero\n2.Febrero\n3.Marzo\n4.Abril\n5.Mayo\n6.Junio\n7.Julio\n8.Agosto\n9.Septiembre\n10.Octubre\n11.Noviembre\n12.Diciembre");
                _cells.GetCell(15, TitleRow02).Value = "START YEAR";

                _cells.GetCell(ResultColumn02, TitleRow02).Value = "RESULTADO";
                _cells.GetCell(ResultColumn02, TitleRow02).Style = _cells.GetStyle(StyleConstants.TitleResult);

                //Adicionar validaciones
                _cells.SetValidationList(_cells.GetCell("B3"), ValidationSheetName, 1);

                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat =
                    NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja: " + ex.Message);
            }
        }

        public void ReviewMstList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearTableRange(TableName01);

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value2);
                var workGroup = _cells.GetEmptyIfNull(_cells.GetCell("B4").Value2);
                var schedIndicator = _cells.GetEmptyIfNull(_cells.GetCell("B5").Value2);

                var listmst = MstActions.FetchMaintenanceScheduleTask(_eFunctions, districtCode, workGroup, null, null, null, null, schedIndicator);
                var i = TitleRow01 + 1;
                foreach (var mst in listmst)
                {
                    try
                    {
                        _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                        _cells.GetCell("B3").Value2 = "" + mst.DistrictCode;
                        _cells.GetCell(1, i).Value2 = "" + mst.WorkGroup;
                        _cells.GetCell(2, i).Value2 = "" + mst.RecType;
                        _cells.GetCell(3, i).Value2 = mst.RecType == MstType.Equipment
                            ? "" + mst.EquipmentNo
                            : "" + mst.EquipmentGrpId;
                        _cells.GetCell(4, i).Value2 = "" + mst.CompCode;
                        _cells.GetCell(5, i).Value2 = "" + mst.CompModCode;
                        _cells.GetCell(6, i).Value2 = "" + mst.MaintenanceSchTask;
                        _cells.GetCell(7, i).Value2 = "" + mst.SchedDescription1;
                        _cells.GetCell(8, i).Value2 = "" + mst.SchedDescription2;

                        _cells.GetCell(9, i).Value2 = "" + mst.StdJobNo;

                        _cells.GetCell(10, i).Value2 = "" + mst.JobDescCode;
                        _cells.GetCell(11, i).Value2 = "" + mst.AssignPerson;
                        _cells.GetCell(12, i).Value2 = "" + mst.SchedInd;

                        _cells.GetCell(13, i).Value2 = "" + mst.SchedFreq1;
                        _cells.GetCell(14, i).Value2 = "" + mst.StatType1;

                        _cells.GetCell(15, i).Value2 = "" + mst.LastSchedStat1;
                        _cells.GetCell(16, i).Value2 = "" + mst.LastPerfStat1;
                        _cells.GetCell(17, i).Value2 = "" + mst.LastSchedDate;
                        _cells.GetCell(18, i).Value2 = "" + mst.LastPerfDate;

                        _cells.GetCell(19, i).Value2 = "" + mst.NextSchedDate;
                        _cells.GetCell(20, i).Value2 = "" + mst.NextSchedStat;

                        _cells.GetCell(25, i).Value2 = "" + mst.OccurrenceType;

                        _cells.GetCell(26, i).Value2 = "" + mst.DayOfWeek;
                        _cells.GetCell(27, i).Value2 = "" + mst.DayOfMonth;

                        _cells.GetCell(28, i).Value2 = "" + mst.StartMonth;
                        _cells.GetCell(29, i).Value2 = "" + mst.StartYear;

                        string freqDescription;
                        if (mst.SchedInd.Equals("1")) //Last Schedule Date
                            freqDescription = mst.SchedFreq1 + "Days/LSD";
                        else if (mst.SchedInd.Equals("2")) //Last Schedule Stat
                            freqDescription = mst.SchedFreq1 + mst.StatType1 + "/LSS";
                        else if (mst.SchedInd.Equals("3")) //Last Performed Date
                            freqDescription = mst.SchedFreq1 + "Days/LPD";
                        else if (mst.SchedInd.Equals("4")) //Last Performed Stat
                            freqDescription = mst.SchedFreq1 + mst.StatType1 + "/LPS";
                        else if (mst.SchedInd.Equals("7")) //Fixed Date
                            freqDescription = "Day " + mst.DayOfMonth + "/" + mst.SchedFreq1 + "Months";
                        else if (mst.SchedInd.Equals("8")) //Fixed Day
                        {
                            string shortDoW = Enum.GetName(typeof(DayOfWeek), Convert.ToInt32(mst.DayOfWeek) % 7);
                            shortDoW = !string.IsNullOrWhiteSpace(shortDoW) ? shortDoW.Substring(0, 3) : "";
                            freqDescription = MyMath.ToOrdinal(Convert.ToInt16(mst.OccurrenceType)) + " " + shortDoW + "/" +
                                                mst.SchedFreq1 + "Months";
                        }
                        else if (mst.SchedInd.Equals("9")) //Inactive
                            freqDescription = "INACTIVE";
                        else
                            freqDescription = "ERROR";
                        _cells.GetCell(21, i).Value2 = "" + freqDescription;
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:GetMstList()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(1, i).Select();
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
            }

        }
        public void ReReviewMstList()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var districtCode = _cells.GetEmptyIfNull(_cells.GetCell("B3").Value2);

                var i = TitleRow01 + 1;
                while (!string.IsNullOrEmpty("" + _cells.GetCell(3, i).Value))
                {
                    try
                    {
                        _cells.GetRange(1, i, ResultColumn01, i).Style = StyleConstants.Normal;
                        var workGroup = "" + _cells.GetCell(1, i).Value2;
                        var equipmentNo = "" + _cells.GetCell(3, i).Value2;
                        var compCode = "" + _cells.GetCell(4, i).Value2;
                        var compModCode = "" + _cells.GetCell(5, i).Value2;
                        var taskNo = "" + _cells.GetCell(6, i).Value2;

                        var mst = MstActions.FetchMaintenanceScheduleTask(_eFunctions, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo);


                        _cells.GetCell("B3").Value2 = "" + mst.DistrictCode;
                        _cells.GetCell(1, i).Value2 = "" + mst.WorkGroup;
                        _cells.GetCell(2, i).Value2 = "" + mst.RecType;
                        _cells.GetCell(3, i).Value2 = "'" + (mst.RecType == MstType.Equipment ? "" + mst.EquipmentNo : "" + mst.EquipmentGrpId);
                        _cells.GetCell(4, i).Value2 = "" + mst.CompCode;
                        _cells.GetCell(5, i).Value2 = "" + mst.CompModCode;
                        _cells.GetCell(6, i).Value2 = "" + mst.MaintenanceSchTask;
                        _cells.GetCell(7, i).Value2 = "" + mst.SchedDescription1;
                        _cells.GetCell(8, i).Value2 = "" + mst.SchedDescription2;
                        _cells.GetCell(9, i).Value2 = "" + mst.StdJobNo;
                        _cells.GetCell(10, i).Value2 = "" + mst.JobDescCode;
                        _cells.GetCell(11, i).Value2 = "" + mst.AssignPerson;
                        _cells.GetCell(12, i).Value2 = "" + mst.SchedInd;
                        _cells.GetCell(13, i).Value2 = "" + mst.SchedFreq1;
                        _cells.GetCell(14, i).Value2 = "" + mst.StatType1;
                        _cells.GetCell(15, i).Value2 = "" + mst.LastSchedStat1;
                        _cells.GetCell(16, i).Value2 = "" + mst.LastPerfStat1;
                        _cells.GetCell(17, i).Value2 = "" + mst.LastSchedDate;
                        _cells.GetCell(18, i).Value2 = "" + mst.LastPerfDate;
                        _cells.GetCell(19, i).Value2 = "" + mst.NextSchedDate;
                        _cells.GetCell(20, i).Value2 = "" + mst.NextSchedStat;
                        _cells.GetCell(25, i).Value2 = "" + mst.OccurrenceType;
                        _cells.GetCell(26, i).Value2 = "" + mst.DayOfWeek;
                        _cells.GetCell(27, i).Value2 = "" + mst.DayOfMonth;
                        _cells.GetCell(28, i).Value2 = "" + mst.StartMonth;
                        _cells.GetCell(29, i).Value2 = "" + mst.StartYear;

                        string freqDescription;
                        if (mst.SchedInd.Equals("1")) //Last Schedule Date
                            freqDescription = mst.SchedFreq1 + "Days/LSD";
                        else if (mst.SchedInd.Equals("2")) //Last Schedule Stat
                            freqDescription = mst.SchedFreq1 + mst.StatType1 + "/LSS";
                        else if (mst.SchedInd.Equals("3")) //Last Performed Date
                            freqDescription = mst.SchedFreq1 + "Days/LPD";
                        else if (mst.SchedInd.Equals("4")) //Last Performed Stat
                            freqDescription = mst.SchedFreq1 + mst.StatType1 + "/LPS";
                        else if (mst.SchedInd.Equals("7")) //Fixed Date
                            freqDescription = "Day " + mst.DayOfMonth + "/" + mst.SchedFreq1 + "Months";
                        else if (mst.SchedInd.Equals("8")) //Fixed Day
                        {
                            string shortDoW = Enum.GetName(typeof(DayOfWeek), Convert.ToInt32(mst.DayOfWeek) % 7);
                            shortDoW = !string.IsNullOrWhiteSpace(shortDoW) ? shortDoW.Substring(0, 3) : "";
                            freqDescription = MyMath.ToOrdinal(Convert.ToInt16(mst.OccurrenceType)) + " " + shortDoW + "/" +
                                                mst.SchedFreq1 + "Months";
                        }
                        else if (mst.SchedInd.Equals("9")) //Inactive
                            freqDescription = "INACTIVE";
                        else
                            freqDescription = "ERROR";
                        _cells.GetCell(21, i).Value2 = "" + freqDescription;

                        _cells.GetCell(ResultColumn01, i).Value2 = "RECONSULTADA";
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(1, i).Style = StyleConstants.Error;
                        _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                        Debugger.LogError("RibbonEllipse.cs:GetMstList()", ex.Message);
                    }
                    finally
                    {
                        _cells.GetCell(1, i).Select();
                        _eFunctions.CloseConnection();
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Error: " + ex.Message);
            }
            finally
            {
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                if (_cells != null) _cells.SetCursorDefault();
            }

        }

        public void CreateMstList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            //Para Servicios WSDL
            var opSheet = new EllipseMaintSchedTaskClassLibrary.MaintSchedTskService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                maxInstancesSpecified = true,
                returnWarningsSpecified = Debugger.DebugWarnings,
            };
            var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
            //
            //Para Servicios POST
            var urlEnvironment = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label, "POST");
            _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnvironment);
            //

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipseDsct, _frmAuth.EllipsePost);



            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var mst = new MaintenanceScheduleTask();
                    mst.DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value2);
                    mst.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                    mst.EquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals(MstType.Egi)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2)
                        : null;
                    mst.EquipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals(MstType.Equipment)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2)
                        : null;
                    mst.CompCode = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value2);
                    mst.CompModCode = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value2);
                    mst.MaintenanceSchTask = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value2);

                    mst.SchedDescription1 = MyUtilities.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null;
                    mst.SchedDescription2 = MyUtilities.IsTrue(_cells.GetCell(8, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value) : null;

                    mst.StdJobNo = MyUtilities.IsTrue(_cells.GetCell(9, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value) : null;

                    mst.JobDescCode = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null;
                    mst.AssignPerson = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    mst.SchedInd = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;

                    mst.AutoRequisitionInd = "N";
                    mst.MsHistFlag = "Y";

                    mst.StatType1 = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null;

                    mst.LastSchedStat1 = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    mst.LastPerfStat1 = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    //Solo se trabaja con una estadística. Cuando se considere trabajar con dos, hay que reevaluar esta sección
                    mst.StatType2 = null;
                    mst.LastSchedStat2 = null;
                    mst.SchedFreq2 = null;
                    mst.LastPerfStat2 = null;

                    mst.LastSchedDate = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null;
                    mst.LastPerfDate = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;

                    mst.StatutoryFlg = "N";


                    var indicator = Convert.ToInt16(mst.SchedInd);

                    if (indicator == 3 || indicator == 4)
                        mst.AllowMultiple = "N";
                    else
                        mst.AllowMultiple = "Y";

                    var frequency = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    var nextSchedStat = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    var nextSchedValue = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null;
                    var nextSchedDate = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    var occurrenceType = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null;
                    var dayOfWeek = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null;
                    string dayOfMonth = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null;
                    var startMonth = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null;
                    var startYear = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;


                    if (indicator >= 1 && indicator <= 4 || indicator == 9)
                    {
                        mst.SchedFreq1 = frequency;
                        mst.NextSchedStat = nextSchedStat;
                        mst.NextSchedValue = nextSchedValue;
                        mst.NextSchedDate = nextSchedDate;
                        mst.SchedFreq1 = frequency;
                        MstActions.CreateMaintenanceScheduleTaskPost(_eFunctions, mst);
                        //var replySheet = MstActions.CreateMaintenanceScheduleTask(urlService, opSheet, mst);
                    }
                    else if (indicator >= 7 && indicator <= 8)
                    {
                        mst.OccurrenceType = occurrenceType;
                        mst.DayOfWeek = dayOfWeek;
                        mst.DayOfMonth = dayOfMonth;
                        mst.SchedFreq1 = frequency;
                        mst.StartMonth = startMonth?.PadLeft(2, '0') ?? "";
                        mst.StartYear = startYear;

                        MstActions.CreateMaintenanceScheduleTaskPost(_eFunctions, mst);
                        //var replySheet = MstActions.CreateMaintenanceScheduleTask(urlService, opSheet, mst);
                    }
                    else
                    {
                        throw new Exception("Indicador de Programación No Válido");
                    }
                    
                    _cells.GetCell(ResultColumn01, i).Value = "CREADA " + mst.EquipmentNo + " " + mst.MaintenanceSchTask;
                    _cells.GetCell(6, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CreateMstList()", ex.Message);
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

        public void UpdateMstList(bool usePost = false)
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);


            if (usePost)
            {
                var alertResult = MessageBox.Show("Actualizar por post cambiará todos los encabezados de actualización a verdaderos, por lo que se actualizarán todos los datos de la hoja.\n¿Está seguro que desea continuar?", "Alerta de Actualización", MessageBoxButtons.YesNo);
                if(alertResult == DialogResult.Yes)
                    for (var k = 7; k < ResultColumn01; k++)
                        _cells.GetCell(k, TitleRow01 - 1).Value = "true";
                else
                    return;
            }
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;

            //Para Servicios WSDL
            var opSheet = new EllipseMaintSchedTaskClassLibrary.MaintSchedTskService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings,
                maxInstancesSpecified = true,
                returnWarningsSpecified = Debugger.DebugWarnings,
            };
            var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
            //
            //Para Servicios POST
            var urlEnvironment = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label, "POST");
            _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnvironment);
            //

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var mst = new MaintenanceScheduleTask();
                    mst.DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value2);
                    mst.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                    mst.EquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals(MstType.Egi)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2)
                        : null;
                    mst.EquipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals(MstType.Equipment)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2)
                        : null;
                    mst.CompCode = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value2);
                    mst.CompModCode = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value2);
                    mst.MaintenanceSchTask = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value2);

                    mst.SchedDescription1 = MyUtilities.IsTrue(_cells.GetCell(7, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value) : null;
                    mst.SchedDescription2 = MyUtilities.IsTrue(_cells.GetCell(8, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value) : null;

                    mst.StdJobNo = MyUtilities.IsTrue(_cells.GetCell(9, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value) : null;

                    mst.JobDescCode = MyUtilities.IsTrue(_cells.GetCell(10, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value) : null;
                    mst.AssignPerson = MyUtilities.IsTrue(_cells.GetCell(11, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value) : null;
                    mst.SchedInd = MyUtilities.IsTrue(_cells.GetCell(12, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value) : null;

                    mst.AutoRequisitionInd = "N";
                    mst.MsHistFlag = "Y";

                    mst.StatType1 = MyUtilities.IsTrue(_cells.GetCell(14, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value) : null;

                    mst.LastSchedStat1 = MyUtilities.IsTrue(_cells.GetCell(15, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value) : null;
                    mst.LastPerfStat1 = MyUtilities.IsTrue(_cells.GetCell(16, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value) : null;
                    //Solo se trabaja con una estadística. Cuando se considere trabajar con dos, hay que reevaluar esta sección
                    mst.StatType2 = null;
                    mst.LastSchedStat2 = null;
                    mst.SchedFreq2 = null;
                    mst.LastPerfStat2 = null;

                    mst.LastSchedDate = MyUtilities.IsTrue(_cells.GetCell(17, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value) : null;
                    mst.LastPerfDate = MyUtilities.IsTrue(_cells.GetCell(18, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value) : null;

                    mst.StatutoryFlg = "N";

                    if (mst.SchedInd == "3" || mst.SchedInd == "4")
                        mst.AllowMultiple = "N";
                    else
                        mst.AllowMultiple = "Y";

                    var indicator = Convert.ToInt16(mst.SchedInd);
                    var frequency = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    var nextSchedStat = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    var nextSchedValue = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null;
                    var nextSchedDate = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    var occurrenceType = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null;
                    var dayOfWeek = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null;
                    string dayOfMonth = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null;

                    var startMonth = MyUtilities.IsTrue(_cells.GetCell(28, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(28, i).Value) : null;
                    var startYear = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;

                    if (string.IsNullOrWhiteSpace(mst.SchedInd))//No hay ajuste de indicador
                    {
                        if (usePost)
                            MstActions.ModifyMaintenanceScheduleTaskPost(_eFunctions, mst);
                        else
                            MstActions.ModifyMaintenanceScheduleTask(urlService, opSheet, mst);
                    }
                    else if ( (indicator >= 1 && indicator <= 4) || indicator == 9 ) 
                    {
                        mst.NextSchedStat = nextSchedStat;
                        mst.NextSchedValue = nextSchedValue;
                        mst.NextSchedDate = nextSchedDate;
                        mst.SchedFreq1 = frequency;
                        if(usePost)
                            MstActions.ModifyMaintenanceScheduleTaskPost(_eFunctions, mst);
                        else
                            MstActions.ModifyMaintenanceScheduleTask(urlService, opSheet, mst);
                        
                    }
                    else if (indicator >= 7 && indicator <= 8)
                    {
                        mst.OccurrenceType = occurrenceType;
                        mst.DayOfWeek = dayOfWeek;
                        mst.DayOfMonth = dayOfMonth;
                        mst.SchedFreq1 = frequency;
                        mst.StartMonth = startMonth?.PadLeft(2, '0') ?? "";
                        mst.StartYear = startYear;
                        if(usePost)
                            MstActions.ModifyMaintenanceScheduleTaskPost(_eFunctions, mst);
                        else
                            MstActions.ModifyMaintenanceScheduleTask(urlService, opSheet, mst);
                        
                    }
                    else
                    {
                        throw new Exception("Indicador de Programación No Válido");
                    }

                    _cells.GetCell(ResultColumn01, i).Value = "ACTUALIZADA";
                    _cells.GetCell(6, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:UpdateMstList()", ex.Message);
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

        public void ModifyNextScheduleList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);
            var i = TitleRow02 + 1;
            var opSheet = new EllipseMaintSchedTaskClassLibrary.MaintSchedTskService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            var urlEnvironment = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label, "POST");
            _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlEnvironment);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var mst = new MaintenanceScheduleTask();
                    mst.DistrictCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B3").Value2);
                    mst.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                    mst.EquipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2);
                    mst.CompCode = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2);
                    mst.CompModCode = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value2);
                    mst.MaintenanceSchTask = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value2);
                    mst.SchedInd = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value2);
                    mst.AllowMultiple = "Y"
                    var indicator = Convert.ToInt16(mst.SchedInd);

                    var nextSchedStat = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value2);
                    var nextSchedValue = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value2);
                    var nextSchedDate = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value2);

                    var occurrenceType = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value2);
                    var dayOfWeek = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value2);
                    string dayOfMonth = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value2);
                    var frequency = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value2);
                    var startMonth = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value2);
                    var startYear = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value2);
                    dayOfMonth = dayOfMonth.PadLeft(2, '0');

                    var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);
                    if (indicator >= 1 && indicator <= 4)
                    {
                        mst.NextSchedStat = nextSchedStat;
                        mst.NextSchedValue = nextSchedValue;
                        mst.NextSchedDate = nextSchedDate;

                        MstActions.ModNextSchedMaintenanceScheduleTask(urlService, opSheet, mst);
                    }
                    else if (indicator >= 7 && indicator <= 8)
                    {
                        mst.OccurrenceType = occurrenceType;
                        mst.DayOfWeek = dayOfWeek;
                        mst.DayOfMonth = dayOfMonth;
                        mst.SchedFreq1 = frequency;
                        mst.StartMonth = startMonth;
                        mst.StartYear = startYear;

                        MstActions.ModifyMaintenanceScheduleTaskPost(_eFunctions, mst);
                    }
                    else
                    {
                        throw new Exception("Indicador de Programación No Válido");
                    }
                    _cells.GetCell(ResultColumn02, i).Value = "REPROGRAMADA";
                    _cells.GetCell(5, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(5, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:ModifyNextScheduleList()", ex.Message);
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

        public void DeleteMstList()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;

            var opSheet = new EllipseMaintSchedTaskClassLibrary.MaintSchedTskService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(2, i).Value))
            {
                try
                {
                    var mst = new MaintenanceScheduleTask();
                    mst.DistrictCode = "ICOR";
                    mst.WorkGroup = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2);
                    mst.EquipmentGrpId = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals(MstType.Egi)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2)
                        : null;
                    mst.EquipmentNo = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2).Equals(MstType.Equipment)
                        ? _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2)
                        : null;
                    mst.CompCode = _cells.GetEmptyIfNull(_cells.GetCell(4, i).Value2);
                    mst.CompModCode = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value2);
                    mst.MaintenanceSchTask = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value2);

                    var urlService = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label);

                    MstActions.DeleteMaintenanceScheduleTask(urlService, opSheet, mst);

                    _cells.GetCell(ResultColumn01, i).Value = "ELIMINADA";
                    _cells.GetCell(6, i).Style = StyleConstants.Success;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                }
                catch (Exception ex)
                {
                    _cells.GetCell(1, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:DeleteMstList()", ex.Message);
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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }


    }


}