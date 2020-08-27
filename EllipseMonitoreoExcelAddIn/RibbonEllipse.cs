using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseMonitoreoExcelAddIn.CondMeasurementService;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Connections;
using Excel = Microsoft.Office.Interop.Excel;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using System.Web.Services.Ellipse;

using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
//using Authenticator = EllipseMonitoreoExcelAddIn.AuthenticatorService;
// ReSharper disable AccessToStaticMemberViaDerivedType
namespace EllipseMonitoreoExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Excel.Application _excelApp;
        string SheetName01 = "Monitoreo";
        string ColHeader = "AK";
        string ColFinal = "AL";
        int ColFin = 38;
        string ColOcultar = "AG1";
        int RowCabezera = 7;
        int RowInicial = 8;
        int maxRow = 10000;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
            _excelApp = Globals.ThisAddIn.Application;

            var enviroments = Environments.GetEnvironmentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }
        public void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new FormAuthenticate();

            var defaultConfig = new Settings.Options();
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
        }
        private void formatear_Click(object sender, RibbonControlEventArgs e)
        {
            setSheetHeaderData();
            Centrar();
        }

        public void setSheetHeaderData()
        {
            try
            {
                this._excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                this._excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                if (this._cells == null)

                    this._cells = new ExcelStyleCells(this._excelApp);

                _cells.GetCell(ColFinal + "1").Value = "OBLIGATORIO";
                _cells.GetCell(ColFinal + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ColFinal + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(ColFinal + "2").Value = "OPCIONAL";
                _cells.GetCell(ColFinal + "2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(ColFinal + "3").Value = "INFORMATIVO";
                _cells.GetCell(ColFinal + "3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(ColFinal + "4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell(ColFinal + "4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell(ColFinal + "5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell(ColFinal + "5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange(ColOcultar, "XFD1048576").Columns.Hidden = true;

                _cells.GetCell("A" + RowCabezera).Value = "FECHA";
                _cells.GetCell("A" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("A7").AddComment("MMDDYYYY");

                _cells.GetCell("B" + RowCabezera).Value = "MUESTRA";
                _cells.GetCell("B" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("C" + RowCabezera).Value = "EQUIPO";
                _cells.GetCell("C" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("C7").NumberFormat = "@";

                _cells.GetCell("D" + RowCabezera).Value = "COMPAR";
                _cells.GetCell("D" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("E" + RowCabezera).Value = "HOROM";
                _cells.GetCell("E" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("F" + RowCabezera).Value = "RECHEQ";
                _cells.GetCell("F" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("G" + RowCabezera).Value = "PB";
                _cells.GetCell("G" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("H" + RowCabezera).Value = "CU";
                _cells.GetCell("H" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("I" + RowCabezera).Value = "FE";
                _cells.GetCell("I" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("J" + RowCabezera).Value = "CR";
                _cells.GetCell("J" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("K" + RowCabezera).Value = "AL";
                _cells.GetCell("K" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("L" + RowCabezera).Value = "SI";
                _cells.GetCell("L" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("M" + RowCabezera).Value = "MO";
                _cells.GetCell("M" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("N" + RowCabezera).Value = "NA";
                _cells.GetCell("N" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("O" + RowCabezera).Value = "B";
                _cells.GetCell("O" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("P" + RowCabezera).Value = "HOLLIN";
                _cells.GetCell("P" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("Q" + RowCabezera).Value = "DI";
                _cells.GetCell("Q" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("R" + RowCabezera).Value = "H2";
                _cells.GetCell("R" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("S" + RowCabezera).Value = "VI";
                _cells.GetCell("S" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("T" + RowCabezera).Value = "CAL";
                _cells.GetCell("T" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("U" + RowCabezera).Value = "MG";
                _cells.GetCell("U" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("V" + RowCabezera).Value = "OXIDA";
                _cells.GetCell("V" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("W" + RowCabezera).Value = "NITRA";
                _cells.GetCell("W" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("X" + RowCabezera).Value = "SULFA";
                _cells.GetCell("X" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("Y" + RowCabezera).Value = "P";
                _cells.GetCell("Y" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("Z" + RowCabezera).Value = "ZN";
                _cells.GetCell("Z" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AA" + RowCabezera).Value = "NI";
                _cells.GetCell("AA" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AB" + RowCabezera).Value = "SN";
                _cells.GetCell("AB" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AC" + RowCabezera).Value = "TI";
                _cells.GetCell("AC" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AD" + RowCabezera).Value = "V";
                _cells.GetCell("AD" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AE" + RowCabezera).Value = "CADMIO";
                _cells.GetCell("AE" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AF" + RowCabezera).Value = "BARIO";
                _cells.GetCell("AF" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AG" + RowCabezera).Value = "COMPONENTE";
                _cells.GetCell("AG" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AH" + RowCabezera).Value = "MOD";
                _cells.GetCell("AH" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AI" + RowCabezera).Value = "ZFDM";
                _cells.GetCell("AI" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AJ" + RowCabezera).Value = "ISO>4";
                _cells.GetCell("AJ" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AK" + RowCabezera).Value = "ISO>6";
                _cells.GetCell("AK" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("AL" + RowCabezera).Value = "ISO>14";
                _cells.GetCell("AL" + RowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);


                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MONITOREO - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", ColHeader + "5");
                _cells.GetRange("B1", ColHeader + "5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", ColHeader + "5").Borders.Weight = "2";
                /*Cells.MergeCells("C6", "L11");
                Cells.GetRange("C6", "L11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Cells.GetRange("C6", "L11").Borders.Weight = "2";
                
                */
                _cells.MergeCells("A" + (RowCabezera - 1), ColFinal + (RowCabezera - 1));
                _cells.GetRange("A" + (RowCabezera - 1), ColFinal + (RowCabezera - 1)).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A" + (RowCabezera - 1), ColFinal + (RowCabezera - 1)).Borders.Weight = "2";

                this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _cells.GetCell("A" + RowInicial).Select();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show("Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void Centrar()
        {

            _cells.GetCell("B" + RowInicial + ":" + ColFinal + maxRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            _cells.GetCell("A" + RowInicial + ":" + ColFinal + maxRow).NumberFormat = "@";

        }

        private void cargar_Click(object sender, RibbonControlEventArgs e)
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;

            if (_frmAuth.ShowDialog() == DialogResult.OK)
            {
               /* if (true)
                {
                    frmAuth.EllipseDsct = "ICOR";
                    frmAuth.EllipsePost = "";
                    frmAuth.EllipseUser = "";
                    frmAuth.EllipsePswd = "";*/


                    int CurrentRow = RowInicial;
                    while (!string.IsNullOrEmpty("" + _cells.GetCell("A" + CurrentRow).Value))
                    {

                        for (int i = 7; i <= ColFin; i++)
                        {
                            if (i == 33)
                            {
                                i = i + 2;
                            }
                            string Medicion = "" + _cells.GetCell(i, CurrentRow).Value;
                            Medicion = Medicion.Trim();
                            if (!string.IsNullOrEmpty(Medicion))
                            {
                                string Fecha = "" + _cells.GetCell("A" + CurrentRow).Value;
                                Fecha = Fecha.Substring(4, 4) + Fecha.Substring(0, 2) + Fecha.Substring(2, 2);
                                string Equipo = "" + _cells.GetCell("C" + CurrentRow).Value;
                                Equipo = Equipo.Trim();
                                string TipoMon = "AA";
                                string CompMon = "" + _cells.GetCell("AG" + CurrentRow).Value;
                                CompMon = CompMon.Trim();
                                string ModMon = "" + _cells.GetCell("AH" + CurrentRow).Value;
                                ModMon = ModMon.Trim();
                                string Elemento = "" + _cells.GetCell(i, RowCabezera).Value;
                                Elemento = Elemento.Trim();

                                CondMeasurementService.CondMeasurementService proxySheet = new CondMeasurementService.CondMeasurementService();

                                CondMeasurementService.OperationContext opSheet = new CondMeasurementService.OperationContext();

                                try
                                {
                                    CondMeasurementService.CondMeasurementServiceCreateRequestDTO requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();
                                    CondMeasurementService.CondMeasurementServiceCreateReplyDTO replySheet = new CondMeasurementServiceCreateReplyDTO();

                                    proxySheet.Url = Environments.GetServiceUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                                    opSheet.district = _frmAuth.EllipseDsct;
                                    opSheet.position = _frmAuth.EllipsePost;
                                    opSheet.maxInstances = 100;
                                    opSheet.returnWarnings = Debugger.DebugWarnings;

                                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                    requestParamsSheet.equipmentRef = Equipo;
                                    requestParamsSheet.condMonType = TipoMon;
                                    requestParamsSheet.compCode = CompMon;
                                    requestParamsSheet.compModCode = ModMon;
                                    requestParamsSheet.measureDate = Fecha;
                                    requestParamsSheet.condMonMeas = Elemento;

                                    if (Elemento == "HOLLIN" || Elemento == "OXIDA" || Elemento == "NITRA" || Elemento == "SULFA")
                                    {
                                        //Medicion = Medicion.Replace(".", ",");
                                        requestParamsSheet.measureValue = System.Math.Round(Convert.ToDecimal(Medicion));
                                        requestParamsSheet.measureValueSpecified = true;

                                    }
                                    else
                                    {
                                        requestParamsSheet.measureValue = Convert.ToDecimal(Medicion);
                                        requestParamsSheet.measureValueSpecified = true;
                                    }

                                    replySheet = proxySheet.create(opSheet, requestParamsSheet);

                                    _cells.GetCell(i, CurrentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                                    _cells.GetCell(i, CurrentRow).Select();
                                    this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                                }
                                catch (Exception ex)
                                {
                                    _cells.GetCell(i, CurrentRow).ClearComments();
                                    _cells.GetCell(i, CurrentRow).AddComment(ex.Message);
                                    _cells.GetCell(i, CurrentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                    _cells.GetCell(i, CurrentRow).Select();
                                    this._excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                                    //   ErrorLogger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, EFunctions.debugErrors);
                                }

                            }
                        }

                        CurrentRow++;
                    }

                    MessageBox.Show("Proceso Finalizado");

               // }
            }
        }

        private void Limpiar()
        {
            _cells.GetCell("A" + RowInicial + ":" + ColFinal + maxRow).Clear();
        }

        private void buttonLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            Limpiar();
        }

        private void borrar_Click(object sender, RibbonControlEventArgs e)
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnviroment.SelectedItem.Label;

            if (_frmAuth.ShowDialog() == DialogResult.OK)
            {
                /*if (true)
                {
                    frmAuth.EllipseDsct = "";
                    frmAuth.EllipsePost = "";
                    frmAuth.EllipseUser = "";
                    frmAuth.EllipsePswd = "";*/


                    int CurrentRow = RowInicial;
                    while (!string.IsNullOrEmpty("" + _cells.GetCell("A" + CurrentRow).Value))
                    {

                        for (int i = 7; i <= ColFin; i++)
                        {
                            /*if (i == 33)
                            {
                                i = i + 2;
                            }*/

                            string Fecha = "" + _cells.GetCell("A" + CurrentRow).Value;
                            Fecha = Fecha.Substring(4, 4) + Fecha.Substring(0, 2) + Fecha.Substring(2, 2);
                            string Equipo = "" + _cells.GetCell("C" + CurrentRow).Value;
                            Equipo = Equipo.Trim();
                            string TipoMon = "AA";
                            string CompMon = "" + _cells.GetCell("AG" + CurrentRow).Value;
                            CompMon = CompMon.Trim();
                            string ModMon = "" + _cells.GetCell("AH" + CurrentRow).Value;
                            ModMon = ModMon.Trim();
                            string Elemento = "" + _cells.GetCell(i, RowCabezera).Value;
                            Elemento = Elemento.Trim();

                            CondMeasurementService.CondMeasurementService proxySheet = new CondMeasurementService.CondMeasurementService();

                            CondMeasurementService.OperationContext opSheet = new CondMeasurementService.OperationContext();

                            try
                            {

                                CondMeasurementService.CondMeasurementServiceDeleteRequestDTO requestParamsSheet = new CondMeasurementServiceDeleteRequestDTO();
                                CondMeasurementService.CondMeasurementServiceDeleteReplyDTO replySheet = new CondMeasurementServiceDeleteReplyDTO();

                                proxySheet.Url = Environments.GetServiceUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                                opSheet.district = _frmAuth.EllipseDsct;
                                opSheet.position = _frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                requestParamsSheet.equipmentRef = Equipo;
                                requestParamsSheet.condMonType = TipoMon;
                                requestParamsSheet.compCode = CompMon;
                                requestParamsSheet.compModCode = ModMon;
                                requestParamsSheet.measureDate = Fecha;
                                requestParamsSheet.condMonMeas = Elemento;

                                replySheet = proxySheet.delete(opSheet, requestParamsSheet);

                                _cells.GetCell(i, CurrentRow).Clear();
                                _cells.GetCell(i, CurrentRow).Select();

                            }
                            catch (Exception ex)
                            {
                                _cells.GetCell(i, CurrentRow).ClearComments();
                                _cells.GetCell(i, CurrentRow).AddComment(ex.Message);
                                _cells.GetCell(i, CurrentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                _cells.GetCell(i, CurrentRow).Clear();
                                _cells.GetCell(i, CurrentRow).Select();
                                // ErrorLogger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, EFunctions.debugErrors);
                            }

                        }

                        CurrentRow++;
                    }

                    MessageBox.Show("Proceso Finalizado");

            //    }
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn("Gustavo Vargas", "").ShowDialog();
        }
        private void drpEnviroment_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }
    }
}

