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
using System.Web.Services.Ellipse;
using System.Web.Services.Ellipse.Post;
using EllipseMstExcelAddIn.MaintSchedTskService;
using System.Threading;
using Util = System.Web.Services.Ellipse.Post.Util;

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
        private const int ResultColumn01 = 31;
        private const int ResultColumn02 = 16;
        private const string TableName01 = "MstTable";
        private const string TableName02 = "ReScheduleTable";
        private const string ValidationSheetName = "ValidationSheetMst";

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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                //si si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(UpdateMstList);

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
                _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
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
                _cells.GetCell(28, TitleRow01).Value = "FREQUENCY";
                _cells.GetCell(28, TitleRow01).AddComment("Indicador 7 y 8. Indica la frecuencia del mes (cada cuántos meses)");
                _cells.GetCell(29, TitleRow01).Value = "START MONTH";
                _cells.GetCell(29, TitleRow01).AddComment("Indicador 7 y 8. Indica a partir de qué mes\n1.Enero\n2.Febrero\n3.Marzo\n4.Abril\n5.Mayo\n6.Junio\n7.Julio\n8.Agosto\n24.Septiembre\n10.Octubre\n11.Noviembre\n12.Diciembre");
                _cells.GetCell(30, TitleRow01).Value = "START YEAR";




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

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

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
                        _cells.GetCell(28, i).Value2 = "" + mst.SchedFreq1;

                        _cells.GetCell(29, i).Value2 = "" + mst.StartMonth;
                        _cells.GetCell(30, i).Value2 = "" + mst.StartYear;

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

                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

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
                        _cells.GetCell(28, i).Value2 = "" + mst.SchedFreq1;

                        _cells.GetCell(29, i).Value2 = "" + mst.StartMonth;
                        _cells.GetCell(30, i).Value2 = "" + mst.StartYear;

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
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;

            var opSheet = new OperationContext
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
                    mst.SchedDescription1 = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value2);
                    mst.SchedDescription2 = _cells.GetEmptyIfNull(_cells.GetCell(8, i).Value2);

                    mst.ConAstSegFr = "0";
                    mst.ConAstSegTo = "0";

                    mst.StdJobNo = _cells.GetEmptyIfNull(_cells.GetCell(9, i).Value2);

                    mst.JobDescCode = _cells.GetEmptyIfNull(_cells.GetCell(10, i).Value2);
                    mst.AssignPerson = _cells.GetEmptyIfNull(_cells.GetCell(11, i).Value2);
                    mst.SchedInd = _cells.GetEmptyIfNull(_cells.GetCell(12, i).Value2);

                    mst.AutoRequisitionInd = "N";
                    mst.MsHistFlag = "Y";

                    mst.SchedFreq1 = _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value2);
                    mst.StatType1 = _cells.GetEmptyIfNull(_cells.GetCell(14, i).Value2);

                    mst.LastSchedStat1 = _cells.GetEmptyIfNull(_cells.GetCell(15, i).Value2);
                    mst.LastPerfStat1 = _cells.GetEmptyIfNull(_cells.GetCell(16, i).Value2);
                    //Solo se trabaja con una estadística. Cuando se considere trabajar con dos, hay que reevaluar esta sección
                    mst.StatType2 = null;
                    mst.LastSchedStat2 = null;
                    mst.SchedFreq2 = null;
                    mst.LastPerfStat2 = null;

                    mst.LastSchedDate = _cells.GetEmptyIfNull(_cells.GetCell(17, i).Value2);
                    mst.LastPerfDate = _cells.GetEmptyIfNull(_cells.GetCell(18, i).Value2);

                    mst.StatutoryFlg = "N";


                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
                    var replySheet = MstActions.CreateMaintenanceScheduleTask(urlService, opSheet, mst);

                    _cells.GetCell(ResultColumn01, i).Value = "CREADA " + replySheet.equipmentRef + " " + replySheet.maintenanceSchTask;
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
        public void UpdateMstList()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
            _eFunctions.SetPostService(_frmAuth.EllipseUser, _frmAuth.EllipsePswd, _frmAuth.EllipsePost, _frmAuth.EllipseDsct, urlService);

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;
            const int validationRow = TitleRow01 - 1;
            var opSheet = new OperationContext
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

                    mst.ConAstSegFr = "0";
                    mst.ConAstSegTo = "0";

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

                    var frequency = MyUtilities.IsTrue(_cells.GetCell(13, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(13, i).Value) : null;
                    var nextSchedStat = MyUtilities.IsTrue(_cells.GetCell(22, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(22, i).Value) : null;
                    var nextSchedValue = MyUtilities.IsTrue(_cells.GetCell(23, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(23, i).Value) : null;
                    var nextSchedDate = MyUtilities.IsTrue(_cells.GetCell(24, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(24, i).Value) : null;
                    var occurrenceType = MyUtilities.IsTrue(_cells.GetCell(25, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(25, i).Value) : null;
                    var dayOfWeek = MyUtilities.IsTrue(_cells.GetCell(26, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(26, i).Value) : null;
                    string dayOfMonth = MyUtilities.IsTrue(_cells.GetCell(27, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(27, i).Value) : null;

                    var startMonth = MyUtilities.IsTrue(_cells.GetCell(29, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(29, i).Value) : null;
                    var startYear = MyUtilities.IsTrue(_cells.GetCell(30, validationRow).Value) ? _cells.GetEmptyIfNull(_cells.GetCell(30, i).Value) : null;
                    dayOfMonth = dayOfMonth.PadLeft(2, '0');

                    if (indicator >= 1 && indicator <= 4)
                    {
                        mst.NextSchedStat = nextSchedStat;
                        mst.NextSchedValue = nextSchedValue;
                        mst.NextSchedDate = nextSchedDate;

                        MstActions.ModifyMaintenanceScheduleTaskPost(_eFunctions, mst);
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
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName02, ResultColumn02);
            var i = TitleRow02 + 1;
            var opSheet = new OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                returnWarnings = Debugger.DebugWarnings
            };
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            while (!string.IsNullOrEmpty("" + _cells.GetCell(1, i).Value))
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

                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);
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
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();
            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);
            var i = TitleRow01 + 1;

            var opSheet = new OperationContext
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

                    var urlService = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label);

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

    public static class MstActions
    {
        public static List<MaintenanceScheduleTask> FetchMaintenanceScheduleTask(EllipseFunctions ef, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo, string schedIndicator)
        {
            var sqlQuery = Queries.GetFetchMstListQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo, schedIndicator);
            var mstDataReader = ef.GetQueryResult(sqlQuery);

            var list = new List<MaintenanceScheduleTask>();

            if (mstDataReader == null || mstDataReader.IsClosed || !mstDataReader.HasRows)
            {
                ef.CloseConnection();
                return list;
            }
            while (mstDataReader.Read())
            {
                // ReSharper disable once UseObjectOrCollectionInitializer
                var mst = new MaintenanceScheduleTask();

                mst.DistrictCode = "" + mstDataReader["DSTRCT_CODE"].ToString().Trim();
                mst.WorkGroup = "" + mstDataReader["WORK_GROUP"].ToString().Trim();
                mst.RecType = "" + mstDataReader["REC_700_TYPE"].ToString().Trim();
                mst.EquipmentNo = mst.RecType == MstType.Equipment ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
                mst.EquipmentGrpId = mst.RecType == MstType.Egi ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
                mst.EquipmentDescription = "" + mstDataReader["EQUIPMENT_DESC"].ToString().Trim();
                mst.CompCode = "" + mstDataReader["COMP_CODE"].ToString().Trim();
                mst.CompModCode = "" + mstDataReader["COMP_MOD_CODE"].ToString().Trim();
                mst.MaintenanceSchTask = "" + mstDataReader["MAINT_SCH_TASK"].ToString().Trim();
                mst.JobDescCode = "" + mstDataReader["JOB_DESC_CODE"].ToString().Trim();
                mst.SchedDescription1 = "" + mstDataReader["SCHED_DESC_1"].ToString().Trim();
                mst.SchedDescription2 = "" + mstDataReader["SCHED_DESC_2"].ToString().Trim();
                mst.AssignPerson = "" + mstDataReader["ASSIGN_PERSON"].ToString().Trim();
                mst.StdJobNo = "" + mstDataReader["STD_JOB_NO"].ToString().Trim();
                mst.AutoRequisitionInd = "" + mstDataReader["AUTO_REQ_IND"].ToString().Trim();
                mst.MsHistFlag = "" + mstDataReader["MS_HIST_FLG"].ToString().Trim();
                mst.SchedInd = "" + mstDataReader["SCHED_IND_700"].ToString().Trim();
                mst.SchedFreq1 = "" + mstDataReader["SCHED_FREQ_1"].ToString().Trim();
                mst.StatType1 = "" + mstDataReader["STAT_TYPE_1"].ToString().Trim();
                mst.LastSchedStat1 = "" + mstDataReader["LAST_SCH_ST_1"].ToString().Trim();
                mst.LastPerfStat1 = "" + mstDataReader["LAST_PERF_ST_1"].ToString().Trim();
                mst.SchedFreq2 = "" + mstDataReader["SCHED_FREQ_2"].ToString().Trim();
                mst.StatType2 = "" + mstDataReader["STAT_TYPE_2"].ToString().Trim();
                mst.LastSchedStat2 = "" + mstDataReader["LAST_SCH_ST_2"].ToString().Trim();
                mst.LastPerfStat2 = "" + mstDataReader["LAST_PERF_ST_2"].ToString().Trim();
                mst.LastSchedDate = "" + mstDataReader["LAST_SCH_DATE"].ToString().Trim();
                mst.LastPerfDate = "" + mstDataReader["LAST_PERF_DATE"].ToString().Trim();
                mst.NextSchedDate = "" + mstDataReader["NEXT_SCH_DATE"].ToString().Trim();
                mst.NextSchedStat = "" + mstDataReader["NEXT_SCH_STAT"].ToString().Trim();
                mst.NextSchedValue = "" + mstDataReader["NEXT_SCH_VALUE"].ToString().Trim();
                mst.ShutdownType = "" + mstDataReader["SHUTDOWN_TYPE"].ToString().Trim();
                mst.ShutdownEquip = "" + mstDataReader["SHUTDOWN_EQUIP"].ToString().Trim();
                mst.ShutdownNo = "" + mstDataReader["SHUTDOWN_NO"].ToString().Trim();
                mst.CondMonPos = "" + mstDataReader["COND_MON_POS"].ToString().Trim();
                mst.CondMonType = "" + mstDataReader["COND_MON_TYPE"].ToString().Trim();
                mst.StatutoryFlg = "" + mstDataReader["STATUTORY_FLG"].ToString().Trim();
                mst.OccurrenceType = "" + mstDataReader["OCCURENCE_TYPE"].ToString().Trim();
                mst.DayOfWeek = "" + mstDataReader["DAY_WEEK"].ToString().Trim();
                mst.DayOfMonth = "" + mstDataReader["DAY_MONTH"].ToString().Trim();
                mst.StartYear = "" + mstDataReader["START_YEAR"].ToString().Trim();
                mst.StartMonth = "" + mstDataReader["START_MONTH"].ToString().Trim();
                list.Add(mst);
            }

            return list;
        }
        public static MaintenanceScheduleTask FetchMaintenanceScheduleTask(EllipseFunctions ef, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo)
        {
            var sqlQuery = Queries.GetFetchMstListQuery(ef.dbReference, ef.dbLink, districtCode, workGroup, equipmentNo, compCode, compModCode, taskNo);
            var mstDataReader = ef.GetQueryResult(sqlQuery);

            if (mstDataReader == null || mstDataReader.IsClosed || !mstDataReader.HasRows || !mstDataReader.Read())
                return null;


            // ReSharper disable once UseObjectOrCollectionInitializer
            var mst = new MaintenanceScheduleTask();

            mst.DistrictCode = "" + mstDataReader["DSTRCT_CODE"].ToString().Trim();
            mst.WorkGroup = "" + mstDataReader["WORK_GROUP"].ToString().Trim();
            mst.RecType = "" + mstDataReader["REC_700_TYPE"].ToString().Trim();
            mst.EquipmentNo = mst.RecType == MstType.Equipment ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
            mst.EquipmentGrpId = mst.RecType == MstType.Egi ? "" + mstDataReader["EQUIP_NO"].ToString().Trim() : null;
            mst.EquipmentDescription = "" + mstDataReader["EQUIPMENT_DESC"].ToString().Trim();
            mst.CompCode = "" + mstDataReader["COMP_CODE"].ToString().Trim();
            mst.CompModCode = "" + mstDataReader["COMP_MOD_CODE"].ToString().Trim();
            mst.MaintenanceSchTask = "" + mstDataReader["MAINT_SCH_TASK"].ToString().Trim();
            mst.JobDescCode = "" + mstDataReader["JOB_DESC_CODE"].ToString().Trim();
            mst.SchedDescription1 = "" + mstDataReader["SCHED_DESC_1"].ToString().Trim();
            mst.SchedDescription2 = "" + mstDataReader["SCHED_DESC_2"].ToString().Trim();
            mst.AssignPerson = "" + mstDataReader["ASSIGN_PERSON"].ToString().Trim();
            mst.StdJobNo = "" + mstDataReader["STD_JOB_NO"].ToString().Trim();
            mst.AutoRequisitionInd = "" + mstDataReader["AUTO_REQ_IND"].ToString().Trim();
            mst.MsHistFlag = "" + mstDataReader["MS_HIST_FLG"].ToString().Trim();
            mst.SchedInd = "" + mstDataReader["SCHED_IND_700"].ToString().Trim();
            mst.SchedFreq1 = "" + mstDataReader["SCHED_FREQ_1"].ToString().Trim();
            mst.StatType1 = "" + mstDataReader["STAT_TYPE_1"].ToString().Trim();
            mst.LastSchedStat1 = "" + mstDataReader["LAST_SCH_ST_1"].ToString().Trim();
            mst.LastPerfStat1 = "" + mstDataReader["LAST_PERF_ST_1"].ToString().Trim();
            mst.SchedFreq2 = "" + mstDataReader["SCHED_FREQ_2"].ToString().Trim();
            mst.StatType2 = "" + mstDataReader["STAT_TYPE_2"].ToString().Trim();
            mst.LastSchedStat2 = "" + mstDataReader["LAST_SCH_ST_2"].ToString().Trim();
            mst.LastPerfStat2 = "" + mstDataReader["LAST_PERF_ST_2"].ToString().Trim();
            mst.LastSchedDate = "" + mstDataReader["LAST_SCH_DATE"].ToString().Trim();
            mst.LastPerfDate = "" + mstDataReader["LAST_PERF_DATE"].ToString().Trim();
            mst.NextSchedDate = "" + mstDataReader["NEXT_SCH_DATE"].ToString().Trim();
            mst.NextSchedStat = "" + mstDataReader["NEXT_SCH_STAT"].ToString().Trim();
            mst.NextSchedValue = "" + mstDataReader["NEXT_SCH_VALUE"].ToString().Trim();
            mst.ShutdownType = "" + mstDataReader["SHUTDOWN_TYPE"].ToString().Trim();
            mst.ShutdownEquip = "" + mstDataReader["SHUTDOWN_EQUIP"].ToString().Trim();
            mst.ShutdownNo = "" + mstDataReader["SHUTDOWN_NO"].ToString().Trim();
            mst.CondMonPos = "" + mstDataReader["COND_MON_POS"].ToString().Trim();
            mst.CondMonType = "" + mstDataReader["COND_MON_TYPE"].ToString().Trim();
            mst.StatutoryFlg = "" + mstDataReader["STATUTORY_FLG"].ToString().Trim();
            mst.OccurrenceType = "" + mstDataReader["OCCURENCE_TYPE"].ToString().Trim();
            mst.DayOfWeek = "" + mstDataReader["DAY_WEEK"].ToString().Trim();
            mst.DayOfMonth = "" + mstDataReader["DAY_MONTH"].ToString().Trim();
            mst.StartYear = "" + mstDataReader["START_YEAR"].ToString().Trim();
            mst.StartMonth = "" + mstDataReader["START_MONTH"].ToString().Trim();
            return mst;
        }



        public static MaintSchedTskServiceCreateReplyDTO CreateMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var proxyEquip = new MaintSchedTskService.MaintSchedTskService();
            var request = new MaintSchedTskServiceCreateRequestDTO
            {
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedDescription1 = mst.SchedDescription1,
                schedDescription2 = mst.SchedDescription2,
                workGroup = mst.WorkGroup,
                assignPerson = mst.AssignPerson,
                jobDescCode = mst.JobDescCode,
                stdJobNo = mst.StdJobNo,
                districtCode = mst.DistrictCode,
                autoRequisitionInd = MyUtilities.IsTrue(mst.AutoRequisitionInd),
                autoRequisitionIndSpecified = mst.AutoRequisitionInd != null,
                MSHistFlag = MyUtilities.IsTrue(mst.MsHistFlag),
                MSHistFlagSpecified = mst.MsHistFlag != null,
                schedInd = mst.SchedInd,
                statType1 = mst.StatType1,
                lastSchedStat1 = !string.IsNullOrWhiteSpace(mst.LastSchedStat1)
                    ? Convert.ToDecimal(mst.LastSchedStat1)
                    : 0,
                lastSchedStat1Specified = mst.LastSchedStat1 != null,
                schedFreq1 = !string.IsNullOrWhiteSpace(mst.SchedFreq1)
                    ? Convert.ToDecimal(mst.SchedFreq1)
                    : 0,
                schedFreq1Specified = mst.SchedFreq1 != null,
                lastPerfStat1 = !string.IsNullOrWhiteSpace(mst.LastPerfStat1)
                    ? Convert.ToDecimal(mst.LastPerfStat1)
                    : 0,
                lastPerfStat1Specified = mst.LastPerfStat1 != null,
                statType2 = mst.StatType2,
                lastSchedStat2 = !string.IsNullOrWhiteSpace(mst.LastSchedStat2)
                    ? Convert.ToDecimal(mst.LastSchedStat2)
                    : 0,
                lastSchedStat2Specified = mst.LastSchedStat2 != null,
                schedFreq2 = !string.IsNullOrWhiteSpace(mst.SchedFreq2)
                    ? Convert.ToDecimal(mst.SchedFreq2)
                    : 0,
                schedFreq2Specified = mst.SchedFreq2 != null,
                lastPerfStat2 = !string.IsNullOrWhiteSpace(mst.LastPerfStat2)
                    ? Convert.ToDecimal(mst.LastPerfStat2)
                    : 0,
                lastPerfStat2Specified = mst.LastPerfStat2 != null,
                lastSchedDate = mst.LastSchedDate,
                lastPerfDate = mst.LastPerfDate,
                statutoryFlg = MyUtilities.IsTrue(mst.StatutoryFlg),
                statutoryFlgSpecified = mst.StatutoryFlg != null,
                occurenceType = mst.OccurrenceType,
                dayOfWeek = mst.DayOfWeek,
                dayOfMonth = mst.DayOfMonth,
                startMonth = mst.StartMonth,
                startYear = mst.StartYear,
                conAstSegFrSpecified = true,
                conAstSegFr = 0,
                conAstSegToSpecified = true,
                conAstSegTo = 0

            };

            proxyEquip.Url = urlService + "/MaintSchedTskService";

            return proxyEquip.create(opContext, request);
        }

        public static MaintSchedTskServiceModifyReplyDTO ModifyMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var proxyEquip = new MaintSchedTskService.MaintSchedTskService();
            var request = new MaintSchedTskServiceModifyRequestDTO
            {

                equipmentGrpId = mst.EquipmentGrpId,
                equipmentNo = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedDescription1 = mst.SchedDescription1,
                schedDescription2 = mst.SchedDescription2,
                workGroup = mst.WorkGroup,
                assignPerson = mst.AssignPerson,
                jobDescCode = mst.JobDescCode,
                stdJobNo = mst.StdJobNo,
                districtCode = mst.DistrictCode,
                autoRequisitionInd = MyUtilities.IsTrue(mst.AutoRequisitionInd),
                autoRequisitionIndSpecified = mst.AutoRequisitionInd != null,
                MSHistFlag = MyUtilities.IsTrue(mst.MsHistFlag),
                MSHistFlagSpecified = mst.MsHistFlag != null,
                schedInd = mst.SchedInd,
                statType1 = mst.StatType1,
                lastSchedStat1 = !string.IsNullOrWhiteSpace(mst.LastSchedStat1)
                    ? Convert.ToDecimal(mst.LastSchedStat1)
                    : 0,
                lastSchedStat1Specified = mst.LastSchedStat1 != null,
                schedFreq1 = !string.IsNullOrWhiteSpace(mst.SchedFreq1)
                    ? Convert.ToDecimal(mst.SchedFreq1)
                    : 0,
                schedFreq1Specified = mst.SchedFreq1 != null,
                lastPerfStat1 = !string.IsNullOrWhiteSpace(mst.LastPerfStat1)
                    ? Convert.ToDecimal(mst.LastPerfStat1)
                    : 0,
                lastPerfStat1Specified = mst.LastPerfStat1 != null,
                statType2 = mst.StatType2,
                lastSchedStat2 = !string.IsNullOrWhiteSpace(mst.LastSchedStat2)
                    ? Convert.ToDecimal(mst.LastSchedStat2)
                    : 0,
                lastSchedStat2Specified = mst.LastSchedStat2 != null,
                schedFreq2 = !string.IsNullOrWhiteSpace(mst.SchedFreq2)
                    ? Convert.ToDecimal(mst.SchedFreq2)
                    : 0,
                schedFreq2Specified = mst.SchedFreq2 != null,
                lastPerfStat2 = !string.IsNullOrWhiteSpace(mst.LastPerfStat2)
                    ? Convert.ToDecimal(mst.LastPerfStat2)
                    : 0,
                lastPerfStat2Specified = mst.LastPerfStat2 != null,
                lastSchedDate = mst.LastSchedDate,
                lastPerfDate = mst.LastPerfDate,
                statutoryFlg = MyUtilities.IsTrue(mst.StatutoryFlg),
                statutoryFlgSpecified = mst.StatutoryFlg != null,
                occurenceType = mst.OccurrenceType,
                dayOfWeek = mst.DayOfWeek,
                dayOfMonth = mst.DayOfMonth,
                startMonth = mst.StartMonth,
                startYear = mst.StartYear,
                conAstSegFrSpecified = true,
                conAstSegFr = 1,
                conAstSegToSpecified = true,
                conAstSegTo = 1
            };

            proxyEquip.Url = urlService + "/MaintSchedTskService";
            return proxyEquip.modify(opContext, request);
        }

        public static void ModifyMaintenanceScheduleTaskPost(EllipseFunctions ef, MaintenanceScheduleTask mst)
        {
            ef.InitiatePostConnection();

            var requestXml = "";

            requestXml = requestXml + "<interaction>";
            requestXml = requestXml + "	<actions>";
            requestXml = requestXml + "		<action>";
            requestXml = requestXml + "			<name>service</name>";
            requestXml = requestXml + "			<data>";
            requestXml = requestXml + "				<name>com.mincom.ellipse.service.m8mwp.mst.MSTService</name>";
            requestXml = requestXml + "				<operation>update</operation>";
            requestXml = requestXml + "				<className>mfui.actions.detail::UpdateAction</className>";
            requestXml = requestXml + "				<returnWarnings>true</returnWarnings>";
            requestXml = requestXml + "				<dto    uuid=\"" + Util.GetNewOperationId() + "\" deleted=\"true\" modified=\"false\">";
            requestXml = requestXml + "					<allowMultiple>Y</allowMultiple>";
            requestXml = requestXml + "					<conAstSegFr>" + mst.ConAstSegFr + "</conAstSegFr>";
            requestXml = requestXml + "					<conAstSegFrNumeric>" + mst.ConAstSegFr + "</conAstSegFrNumeric>";
            requestXml = requestXml + "					<conAstSegTo>" + mst.ConAstSegTo + "</conAstSegTo>";
            requestXml = requestXml + "					<conAstSegToNumeric>" + mst.ConAstSegTo + "</conAstSegToNumeric>";
            requestXml = requestXml + "					<dayMonth>" + mst.DayOfMonth + "</dayMonth>";
            requestXml = requestXml + "					<dayWeek> " + mst.DayOfWeek + " </dayWeek>";
            requestXml = requestXml + "					<dstrctCode>" + mst.DistrictCode + "</dstrctCode>";
            requestXml = requestXml + "					<equipEntity>" + mst.EquipmentNo + "</equipEntity>";
            requestXml = requestXml + "					<equipNo>" + mst.EquipmentNo + "</equipNo>";
            requestXml = requestXml + "					<equipRef>" + mst.EquipmentNo + "</equipRef>";
            requestXml = requestXml + "					<fixedScheduling>Y</fixedScheduling>";
            requestXml = requestXml + "					<isInSeries>Y</isInSeries>";
            requestXml = requestXml + "					<isInSuppressionSeries>Y</isInSuppressionSeries>";
            requestXml = requestXml + "					<jobDescCode>" + mst.JobDescCode + "</jobDescCode>";
            requestXml = requestXml + "					<lastPerfDate>" + mst.LastPerfDate + "</lastPerfDate>";
            requestXml = requestXml + "					<lastPerfStat1>" + mst.LastPerfStat1 + "</lastPerfStat1>";
            requestXml = requestXml + "					<lastSchDate>" + mst.LastSchedDate + "</lastSchDate>";
            requestXml = requestXml + "					<lastSchStat1>" + mst.LastSchedStat1 + "</lastSchStat1>";
            requestXml = requestXml + "					<linkedInd>N</linkedInd>";
            requestXml = requestXml + "					<maintSchTask>" + mst.MaintenanceSchTask + "</maintSchTask>";
            requestXml = requestXml + "					<msHistFlg>Y</msHistFlg>";
            requestXml = requestXml + "					<nextSchDate>" + mst.NextSchedDate + "</nextSchDate>";
            requestXml = requestXml + "					<rec700Type>" + mst.RecType + "</rec700Type>";
            requestXml = requestXml + "					<recallTimeHrs>0.00</recallTimeHrs>";
            requestXml = requestXml + "					<schedDesc1>" + mst.SchedDescription1 + "</schedDesc1>";
            requestXml = requestXml + "					<schedFreq1>" + mst.SchedFreq1 + "</schedFreq1>";
            requestXml = requestXml + "					<schedInd700>" + mst.SchedInd + "</schedInd700>";
            requestXml = requestXml + "					<startMonth>" + mst.StartMonth + "</startMonth>";
            requestXml = requestXml + "					<startYear>" + mst.StartYear + "</startYear>";
            requestXml = requestXml + "					<workGroup>" + mst.WorkGroup + "</workGroup>";
            requestXml = requestXml + "					<autoReqInd>N</autoReqInd>";
            requestXml = requestXml + "					<statType1>" + mst.StatType1 + "</statType1>";
            requestXml = requestXml + "					<statType2>" + mst.StatType2 + "</statType2>";
            requestXml = requestXml + "					<nextSchStat>" + mst.NextSchedStat + "</nextSchStat>";
            requestXml = requestXml + "					<nextSchValue>" + mst.NextSchedValue + "</nextSchValue>";
            requestXml = requestXml + "					<statutoryFlg>N</statutoryFlg>";
            requestXml = requestXml + "					<hideSuppressed>Y</hideSuppressed>";
            requestXml = requestXml + "				</dto>";
            requestXml = requestXml + "			</data>";
            requestXml = requestXml + "			<id>" + Util.GetNewOperationId() + "</id>";
            requestXml = requestXml + "		</action>";
            requestXml = requestXml + "	</actions>";
            requestXml = requestXml + "	<chains/>";
            requestXml = requestXml + "	<connectionId>" + ef.PostServiceProxy.ConnectionId + "</connectionId>";
            requestXml = requestXml + "	<application>msemst</application>";
            requestXml = requestXml + "	<applicationPage>read</applicationPage>";
            requestXml = requestXml + "	<transaction>true</transaction>";
            requestXml = requestXml + "</interaction>";

            requestXml = requestXml.Replace("&", "&amp;");
            var responseDto = ef.ExecutePostRequest(requestXml);

            if (!responseDto.GotErrorMessages()) return;
            var errorMessage = responseDto.Errors.Aggregate("", (current, msg) => current + (msg.Field + " " + msg.Text));
            if (!errorMessage.Equals(""))
                throw new Exception(errorMessage);
        }

        public static MaintSchedTskServiceModNextSchedReplyDTO ModNextSchedMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {

            var proxyEquip = new MaintSchedTskService.MaintSchedTskService();
            var request = new MaintSchedTskServiceModNextSchedRequestDTO()
            {
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                nextSchedDate = mst.NextSchedDate,
                nextSchedValSpecified = string.IsNullOrWhiteSpace(mst.NextSchedValue),
                nextStat = mst.NextSchedStat,
                nextSchedVal = Convert.ToDecimal(mst.NextSchedValue)
            };

            proxyEquip.Url = urlService + "/MaintSchedTskService";
            return proxyEquip.modNextSched(opContext, request);
        }

        public static MaintSchedTskServiceDeleteReplyDTO DeleteMaintenanceScheduleTask(string urlService, OperationContext opContext, MaintenanceScheduleTask mst)
        {
            var proxyEquip = new MaintSchedTskService.MaintSchedTskService { Url = urlService + "/MaintSchedTskService" };

            //actualizamos primero el indicador y eliminamos la frecuencia
            var requestUpdate = new MaintSchedTskServiceModifyRequestDTO
            {
                workGroup = mst.WorkGroup,
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask,
                schedFreq1 = 0,
                schedFreq2 = 0,
                schedInd = "9",
                schedFreq1Specified = true,
                schedFreq2Specified = true,
                statType1 = "",
                statType2 = ""
            };


            var request = new MaintSchedTskServiceDeleteRequestDTO
            {
                equipmentGrpId = mst.EquipmentGrpId,
                equipmentRef = mst.EquipmentNo,
                compCode = mst.CompCode,
                compModCode = mst.CompModCode,
                maintenanceSchTask = mst.MaintenanceSchTask
            };

            proxyEquip.modify(opContext, requestUpdate);
            return proxyEquip.delete(opContext, request);
        }

        public static class Queries
        {
            public static string GetFetchMstListQuery(string dbReference, string dbLink, string districtCode, string workGroup, string equipmentNo, string compCode, string compModCode, string taskNo, string schedIndicator = null)
            {
                if (!string.IsNullOrWhiteSpace(districtCode))
                    districtCode = " AND MST.DSTRCT_CODE = '" + districtCode + "'";
                if (!string.IsNullOrWhiteSpace(workGroup))
                    workGroup = " AND MST.WORK_GROUP = '" + workGroup + "'";
                if (!string.IsNullOrWhiteSpace(equipmentNo))
                    equipmentNo = " AND MST.EQUIP_NO = '" + equipmentNo + "'";
                if (!string.IsNullOrWhiteSpace(compCode))
                    compCode = " AND MST.COMP_CODE = '" + compCode + "'";
                if (!string.IsNullOrWhiteSpace(compModCode))
                    compModCode = " AND MST.COMP_MOD_CODE = '" + compModCode + "'";
                if (!string.IsNullOrWhiteSpace(taskNo))
                    taskNo = " AND MST.MAINT_SCH_TASK = '" + taskNo + "'";

                //establecemos los parámetros de estado de orden
                schedIndicator = MyUtilities.GetCodeValue(schedIndicator);
                string statusIndicator;
                if (string.IsNullOrEmpty(schedIndicator))
                    statusIndicator = "";
                else if (schedIndicator == MstIndicatorList.Active)
                    statusIndicator = " AND MST.SCHED_IND_700 IN (" + MyUtilities.GetListInSeparator(MstIndicatorList.GetActiveIndicatorCodes(), ",", "'") + ")";
                else if (MstIndicatorList.GetIndicatorNames().Contains(schedIndicator))
                    statusIndicator = " AND MST.SCHED_IND_700 = '" + MstIndicatorList.GetIndicatorCode(schedIndicator) + "'";
                else
                    statusIndicator = "";

                var query = "" +
                               " SELECT" +
                               "     MST.DSTRCT_CODE, MST.WORK_GROUP, MST.REC_700_TYPE, MST.EQUIP_NO, EQ.ITEM_NAME_1 EQUIPMENT_DESC, MST.COMP_CODE, MST.COMP_MOD_CODE, MST.MAINT_SCH_TASK," +
                               "     MST.JOB_DESC_CODE, MST.SCHED_DESC_1, MST.SCHED_DESC_2, MST.ASSIGN_PERSON, MST.STD_JOB_NO, MST.AUTO_REQ_IND, MST.MS_HIST_FLG, MST.SCHED_IND_700," +
                               "     MST.SCHED_FREQ_1, MST.STAT_TYPE_1, MST.LAST_SCH_ST_1, MST.LAST_PERF_ST_1," +
                               "     MST.SCHED_FREQ_2, MST.STAT_TYPE_2, MST.LAST_SCH_ST_2, MST.LAST_PERF_ST_2," +
                               "     MST.LAST_SCH_DATE, MST.LAST_PERF_DATE, MST.NEXT_SCH_DATE, MST.NEXT_SCH_STAT, MST.NEXT_SCH_VALUE," +
                               "     MST.OCCURENCE_TYPE, MST.DAY_WEEK, MST.DAY_MONTH, DECODE(TRIM(MST.LAST_SCH_DATE),NULL,'',SUBSTR(MST.LAST_SCH_DATE,1,4) ) START_YEAR, DECODE(TRIM(MST.LAST_SCH_DATE),NULL,'',SUBSTR(MST.LAST_SCH_DATE,5,2) )START_MONTH, " +
                               "     MST.SHUTDOWN_TYPE , MST.SHUTDOWN_EQUIP, MST.SHUTDOWN_NO, MST.COND_MON_POS, MST.COND_MON_TYPE, MST.STATUTORY_FLG" +
                               " FROM" +
                               "     " + dbReference + ".MSF700" + dbLink + " MST LEFT JOIN " + dbReference + ".MSF600" + dbLink + " EQ ON MST.EQUIP_NO = EQ.EQUIP_NO" +
                               " WHERE" +
                               districtCode +
                               workGroup +
                               equipmentNo +
                               compCode +
                               compModCode +
                               taskNo +
                               statusIndicator +
                               " ORDER BY MST.MAINT_SCH_TASK DESC";
                query = MyUtilities.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

                return query;
            }
        }
    }



}
