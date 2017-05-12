using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseMSE345ExcelAddIn.CondMeasurementService;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using EllipseCommonsClassLibrary;
using EllipseStdTextClassLibrary;
// ReSharper disable FieldCanBeMadeReadOnly.Local

namespace EllipseMSE345ExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;
        private const string SheetName01 = "MSE345";
        private const string SheetName01Mtto = "MSE345_MTTO";

        string _narrativeNo;
        decimal _valor;
        Worksheet _worksheet;
        Microsoft.Office.Tools.Excel.Controls.DateTimePicker _fechaCalendario;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
        }

        private void Crear_Click(object sender, RibbonControlEventArgs e)
        { 
            if(_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01))
            {
                Cargar_Info_Estandar();
            }
            else if(_excelApp.ActiveWorkbook.ActiveSheet.Name.Equals(SheetName01Mtto))
            {
                Cargar_Info_MTTO();
            }
            else
            {
                MessageBox.Show(@"LA HOJA NO CONTIENE EL FORMATO REQUERIDO PARA REALIZAR ESTA ACCION");
            }
        }

        private void Cargar_Info_MTTO()
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

            if (!string.IsNullOrEmpty("" + _cells.GetCell("B7").Value) && !string.IsNullOrEmpty("" + _cells.GetCell("B8").Value) && !string.IsNullOrEmpty("" + _cells.GetCell("B9").Value))
            {

                if (_frmAuth.ShowDialog() == DialogResult.OK)
                {
                    // Cells.getCell("A1").Value = "Conectado";

                    var proxySheet = new CondMeasurementService.CondMeasurementService();

                    var opSheet = new OperationContext();

                    var currentRow = 15;
                    string equipo = "" + _cells.GetCell("B8").Value;
                    string tipoMonitoreo = "" + _cells.GetCell("B7").Value;
                    string fecha = "" + _cells.GetCell("B9").Value;
                    string medida = "" + _cells.GetCell("A" + currentRow).Value;
                    string insp1 = "" + _cells.GetCell("B10").Value;
                    if (!string.IsNullOrEmpty(insp1))
                    {
                        insp1 = insp1.Substring(0, 2);
                    }
                    string insp2 = "" + _cells.GetCell("B11").Value;
                    if (!string.IsNullOrEmpty(insp2))
                    {
                        insp2 = insp2.Substring(0, 2);
                    }
                    string insp3 = "" + _cells.GetCell("B12").Value;
                    if (!string.IsNullOrEmpty(insp3))
                    {
                        insp3 = insp3.Substring(0, 2);
                    }
                    string comentario = "" + _cells.GetCell("H" + currentRow).Value;

                    while (!string.IsNullOrEmpty(medida))
                    {
                        if (string.IsNullOrEmpty(Convert.ToString(_cells.GetCell("G" + currentRow).Value)))
                        {
                            currentRow++;
                            medida = "" + _cells.GetCell("A" + currentRow).Value;
                            comentario = "" + _cells.GetCell("H" + currentRow).Value;
                        }
                        else
                        {
                            _valor = Convert.ToDecimal(_cells.GetCell("G" + currentRow).Value);
                            try
                            {
                                var requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();

                                proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                                opSheet.district = _frmAuth.EllipseDsct;
                                opSheet.position = _frmAuth.EllipsePost;
                                opSheet.maxInstances = 100;
                                opSheet.returnWarnings = Debugger.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                requestParamsSheet.equipmentRef = equipo;
                                requestParamsSheet.condMonType = tipoMonitoreo;
                                requestParamsSheet.measureDate = fecha;
                                requestParamsSheet.condMonMeas = medida;

                                requestParamsSheet.measureValue = Convert.ToDecimal(_valor);
                                requestParamsSheet.measureValueSpecified = true;

                                requestParamsSheet.visInsCode1 = insp1;

                                if (!string.IsNullOrEmpty(insp2))
                                {
                                    requestParamsSheet.visInsCode2 = insp2;
                                }

                                if (!string.IsNullOrEmpty(insp3))
                                {
                                    requestParamsSheet.visInsCode3 = insp3;
                                }

                                proxySheet.create(opSheet, requestParamsSheet);

                                if (!string.IsNullOrEmpty(comentario))
                                {
                                    _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                                    var sqlQuery = "select narrative_no from ellipse.msf345 where substr(99999999999999-rev_meas_data,1,8) = '" + fecha + "' and equip_no = '" + equipo + "' and trim(comp_pos_data) = '" + tipoMonitoreo + "' and trim(cond_mon_meas) = '" + medida + "'";

                                    var odr = _eFunctions.GetQueryResult(sqlQuery);
                                    if (odr.Read())
                                    {
                                        _narrativeNo = odr["narrative_no"] + "";
                                    }

                                    StdText.SetText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, false), "ME" + _narrativeNo, comentario);
                                }

                                _cells.GetCell("I" + currentRow).Value = "OK";
                                _cells.GetCell("I" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                                _cells.GetCell("I" + currentRow).Select();
                            }
                            catch (Exception ex)
                            {
                                _cells.GetCell("I" + currentRow).Value = ex.Message;
                                _cells.GetCell("I" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                                _cells.GetCell("I" + currentRow).Select();
                                Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                            }
                            finally
                            {
                                currentRow++;
                                medida = "" + _cells.GetCell("A" + currentRow).Value;
                                comentario = "" + _cells.GetCell("H" + currentRow).Value;
                            }
                        }
                    }
                    MessageBox.Show(@"Proceso Finalizado Correctamente");
                }
            }
            else
            {
                MessageBox.Show(@"Hay algunos Campos Obligatorios Vacios. Reviselos e Intente Nuevamente");
            }
        }

        private void Cargar_Info_Estandar()
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;
            var monitorType = _cells.GetEmptyIfNull(_cells.GetCell("B6").Value);
            var monitorDate = _cells.GetEmptyIfNull(_cells.GetCell("B7").Value);
            var monitorEquipment = _cells.GetEmptyIfNull(_cells.GetCell("B11").Value);

            if (!string.IsNullOrWhiteSpace(monitorType) && !string.IsNullOrWhiteSpace(monitorDate) && !string.IsNullOrWhiteSpace(monitorEquipment))
            {
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                // Cells.getCell("A1").Value = "Conectado";

                var proxySheet = new CondMeasurementService.CondMeasurementService();

                var opSheet = new OperationContext();

                var currentRow = 14;
                string componentCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("A" + currentRow).Value);
                string modifierCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("B" + currentRow).Value);
                string modifierPosition = _cells.GetNullOrTrimmedValue(_cells.GetCell("C" + currentRow).Value);

                componentCode = componentCode.Equals("-") ? null : componentCode;
                modifierCode = modifierCode.Equals("-") ? null : modifierCode;
                modifierPosition = modifierPosition.Equals("-") ? null : modifierPosition;

                componentCode = componentCode != null && componentCode.Length > 4 && componentCode.Contains("-") ? componentCode.Substring(0, 4) : componentCode;
                modifierCode = modifierCode != null && modifierCode.Length > 4 && modifierCode.Contains("-") ? modifierCode.Substring(0, 4) : modifierCode;
                modifierPosition = modifierPosition != null && modifierPosition.Length > 4 && modifierPosition.Contains("-") ? modifierPosition.Substring(0, 4) : modifierPosition;
                

                    
                string measurementCode = "" + _cells.GetCell("D" + currentRow).Value;
                
                string insp1 = "" + _cells.GetCell("B8").Value;
                if (!string.IsNullOrEmpty(insp1))
                {
                    insp1 = insp1.Substring(0, 2);
                }
                string insp2 = "" + _cells.GetCell("B9").Value;
                if (!string.IsNullOrEmpty(insp2))
                {
                    insp2 = insp2.Substring(0, 2);
                }
                string insp3 = "" + _cells.GetCell("B10").Value;
                if (!string.IsNullOrEmpty(insp3))
                {
                    insp3 = insp3.Substring(0, 2);
                }

                string comentario = "" + _cells.GetCell("K" + currentRow).Value;

                while (!string.IsNullOrEmpty(measurementCode))
                {
                    if (string.IsNullOrEmpty(Convert.ToString(_cells.GetCell("J" + currentRow).Value)))
                    {
                        currentRow++;
                        measurementCode = "" + _cells.GetCell("D" + currentRow).Value;
                        comentario = "" + _cells.GetCell("K" + currentRow).Value;

                        componentCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("A" + currentRow).Value);
                        modifierCode = _cells.GetNullOrTrimmedValue(_cells.GetCell("B" + currentRow).Value);
                        modifierPosition = _cells.GetNullOrTrimmedValue(_cells.GetCell("C" + currentRow).Value);

                        componentCode = componentCode.Equals("-") ? null : componentCode;
                        modifierCode = modifierCode.Equals("-") ? null : modifierCode;
                        modifierPosition = modifierPosition.Equals("-") ? null : modifierPosition;

                        componentCode = componentCode != null && componentCode.Length > 4 && componentCode.Contains("-") ? componentCode.Substring(0, 4) : componentCode;
                        modifierCode = modifierCode != null && modifierCode.Length > 4 && modifierCode.Contains("-") ? modifierCode.Substring(0, 4) : modifierCode;
                        modifierPosition = modifierPosition != null && modifierPosition.Length > 4 && modifierPosition.Contains("-") ? modifierPosition.Substring(0, 4) : modifierPosition;

                    }
                    else
                    {
                        _valor = Convert.ToDecimal(_cells.GetCell("J" + currentRow).Value);
                        try
                        {
                            var requestParamsSheet = new CondMeasurementServiceCreateRequestDTO();

                            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/CondMeasurementService";

                            opSheet.district = _frmAuth.EllipseDsct;
                            opSheet.position = _frmAuth.EllipsePost;
                            opSheet.maxInstances = 100;
                            opSheet.returnWarnings = Debugger.DebugWarnings;

                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            requestParamsSheet.equipmentRef = monitorEquipment;
                            requestParamsSheet.condMonType = monitorType;
                            requestParamsSheet.measureDate = monitorDate;
                            requestParamsSheet.condMonMeas = measurementCode;
                            requestParamsSheet.compCode = componentCode;
                            requestParamsSheet.compModCode = modifierCode;
                            requestParamsSheet.condMonPos = modifierPosition;

                            requestParamsSheet.measureValue = Convert.ToDecimal(_valor);
                            requestParamsSheet.measureValueSpecified = true;

                            requestParamsSheet.visInsCode1 = insp1;

                            if (!string.IsNullOrEmpty(insp2))
                            {
                                requestParamsSheet.visInsCode2 = insp2;
                            }

                            if (!string.IsNullOrEmpty(insp3))
                            {
                                requestParamsSheet.visInsCode3 = insp3;
                            }

                            proxySheet.create(opSheet, requestParamsSheet);

                            if (!string.IsNullOrEmpty(comentario))
                            {
                                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                                var sqlQuery = "select narrative_no from ellipse.msf345 where substr(99999999999999-rev_meas_data,1,8) = '" + monitorDate + "' and equip_no = '" + monitorEquipment + "' and trim(comp_pos_data) = '" + monitorType + "' and trim(cond_mon_meas) = '" + measurementCode + "'";

                                var odr = _eFunctions.GetQueryResult(sqlQuery);
                                if (odr.Read())
                                {
                                    _narrativeNo = odr["narrative_no"] + "";
                                }

                                StdText.SetText(_eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDsct, _frmAuth.EllipsePost, 100, false), "ME" + _narrativeNo, comentario);
                            }

                            _cells.GetCell("L" + currentRow).Value = "OK";
                            _cells.GetCell("L" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                            _cells.GetCell("L" + currentRow).Select();
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell("L" + currentRow).Value = ex.Message;
                            _cells.GetCell("L" + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                            _cells.GetCell("L" + currentRow).Select();
                            Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);

                        }
                        finally
                        {
                            currentRow++;
                            measurementCode = "" + _cells.GetCell("D" + currentRow).Value;
                            comentario = "" + _cells.GetCell("K" + currentRow).Value;
                            componentCode = "" + _cells.GetCell("A" + currentRow).Value;
                            modifierCode = "" + _cells.GetCell("B" + currentRow).Value;
                            modifierPosition = "" + _cells.GetCell("C" + currentRow).Value;

                            componentCode = _cells.GetCell("A" + currentRow).Value != "-" ? componentCode.Substring(0, 4) : "";

                            if (_cells.GetCell("B" + currentRow).Value != "-")
                            {
                                modifierCode = modifierCode.Substring(0, 2);
                                modifierCode = modifierCode.Trim();
                            }
                            else
                            {
                                modifierCode = "";
                            }

                            if (_cells.GetCell("C" + currentRow).Value != "-")
                            {
                                modifierPosition = modifierPosition.Substring(0, 2);
                                modifierPosition = modifierPosition.Trim();
                            }
                            else
                            {
                                modifierPosition = "";
                            }

                        }
                    }
                }
                MessageBox.Show(@"Proceso Finalizado Correctamente");
            }
            else
            {
                MessageBox.Show(@"Hay algunos Campos Obligatorios Vacios. Reviselos e Intente Nuevamente");
            }
        }

        private void Formatear_Click(object sender, RibbonControlEventArgs e)
        {
            SetSheetHeaderData();
            Centrar();
            _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            var groupCells = _worksheet.Range["B11:B11"];
            var groupRange = _worksheet.Controls.AddNamedRange(groupCells, "GroupRange");

            groupRange.Change += changesGroupRange_Change;

            var groupCells2 = _worksheet.Range["B6:B6"];
            var groupRange2 = _worksheet.Controls.AddNamedRange(groupCells2, "GroupRange2");

            groupRange2.Change += changesGroupRange_Change2;

            var groupCells3 = _worksheet.Range["B8:B10"];
            var groupRange3 = _worksheet.Controls.AddNamedRange(groupCells3, "GroupRange3");

            groupRange3.Change += AutoAjuste;
            
        }        

        public void AutoAjuste(Excel.Range target)
        {
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }

        public void SetSheetHeaderData()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                if (_cells == null)

                    _cells = new ExcelStyleCells(_excelApp);

                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange("M1", "XFD1048576").Columns.Hidden = true;

                _cells.GetCell("A13").Value = "COMPONENTE";
                _cells.GetCell("A13").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("B13").Value = "MODIFICADOR";
                _cells.GetCell("B13").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("C13").Value = "POSICION";
                _cells.GetCell("C13").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("D13").Value = "CODIGO";
                _cells.GetCell("D13").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("E13").Value = "DESCRIPCION";
                _cells.GetCell("E13").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("F13").Value = "CAUTION LOW";
                _cells.GetCell("F13").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("G13").Value = "CAUTION";
                _cells.GetCell("G13").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("H13").Value = "DANGER LOW";
                _cells.GetCell("H13").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("I13").Value = "DANGER";
                _cells.GetCell("I13").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("J13").Value = "VALOR ENCONTRADO";
                _cells.GetCell("J13").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                                 
                _cells.GetCell("K13").Value = "COMENTARIO";
                _cells.GetCell("K13").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                                 
                _cells.GetCell("L13").Value = "RESULTADO";
                _cells.GetCell("L13").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "K5");
                _cells.GetRange("B1", "K5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", "K5").Borders.Weight = "2";
                _cells.MergeCells("C6", "L11");
                _cells.GetRange("C6", "L11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("C6", "L11").Borders.Weight = "2";
                _cells.MergeCells("A12", "L12");
                _cells.GetRange("A12", "L12").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A12", "L12").Borders.Weight = "2";

                _cells.GetCell("A6").Value = "MONITOREO";
                _cells.GetCell("A6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationList(_cells.GetCell("B6"), GetMonitoreos());
                _cells.GetCell("B6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B6").Borders.Weight = "2";

                _cells.GetCell("A7").Value = "FECHA";
                _cells.GetCell("A7").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B7").Value = DateTime.Now.ToString("yyyyMMdd");
                _cells.GetCell("B7").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                _cells.GetCell("B7").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B7").Borders.Weight = "2";

                var inspectores = GetInspectores();

                _cells.GetCell("A8").Value = "INSPECTOR 1";
                _cells.GetCell("A8").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B8"), inspectores);
                _cells.GetCell("B8").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B8").Borders.Weight = "2";

                _cells.GetCell("A9").Value = "INSPECTOR 2";
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B9"), inspectores);
                _cells.GetCell("B9").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B9").Borders.Weight = "2";

                _cells.GetCell("A10").Value = "INSPECTOR 3";
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B10"), inspectores);
                _cells.GetCell("B10").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B10").Borders.Weight = "2";

                _cells.GetCell("A11").Value = "EQUIPO";
                _cells.GetCell("A11").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B11").NumberFormat = "@";
                //Cells.setValidationList(Cells.getCell("B11"), getEquipos());
                _cells.GetCell("B11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B11").Borders.Weight = "2";

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _cells.GetCell("B6").Select();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        private void Centrar()
        {
            var row = 6;
            while (row <= 200)
            {
                //Cells.getCell("A" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //Cells.getCell("B" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //Cells.getCell("C" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("D" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //Cells.getCell("E" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("F" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("G" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("H" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("I" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("J" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("K" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("L" + row).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; 
                row++;
            }
        }

        private void CentrarMtto()
        {
                _cells.GetCell("A15:I200").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

        }

        private void Limpiar()
        {
            var row = 14;
            var max = 200;
            while (row <= max)
            {
                _cells.GetCell("A" + row).Clear();
                _cells.GetCell("B" + row).Clear();
                _cells.GetCell("C" + row).Clear();
                _cells.GetCell("D" + row).Clear();
                _cells.GetCell("E" + row).Clear();
                _cells.GetCell("F" + row).Clear();
                _cells.GetCell("G" + row).Clear();
                _cells.GetCell("H" + row).Clear();
                _cells.GetCell("I" + row).Clear();
                _cells.GetCell("J" + row).Clear();
                _cells.GetCell("K" + row).Clear();
                _cells.GetCell("L" + row).Clear();
                row++;
            }
        }

        private void LimpiarMtto()
        {
            var row = 15;
            var max = 200;
            while (row <= max)
            {
                _cells.GetCell("A" + row).Clear();
                _cells.GetCell("B" + row).Clear();
                _cells.GetCell("C" + row).Clear();
                _cells.GetCell("D" + row).Clear();
                _cells.GetCell("E" + row).Clear();
                _cells.GetCell("F" + row).Clear();
                _cells.GetCell("G" + row).Clear();
                _cells.GetCell("H" + row).Clear();
                _cells.GetCell("I" + row).Clear();
                _cells.GetCell("J" + row).Clear();
                _cells.GetCell("K" + row).Clear();
                _cells.GetCell("L" + row).Clear();
                row++;
            }
        }

        public List<string> GetFlotas()
        {
            _eFunctions.SetDBSettings(EnviromentConstants.SigcorProductivo);

            const string sqlQuery = "SELECT DISTINCT TRIM(FLOTA_ELLIPSE) AS FLOTA FROM EQMTLIST WHERE FLOTA_ELLIPSE IS NOT NULL AND ACTIVE_FLG = 'Y' ORDER BY 1";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getFlotas = new List<string>();

            while (odr.Read())
            {
                getFlotas.Add("" + odr["FLOTA"]);
            }
            return getFlotas;
        }

        public List<string> GetEquipos()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var sqlQuery = "select equip_no from ellipse.msf600 where equip_no between '0220701' and '0220999' and equip_no not in ( '02209       ','02208       ') order by 1";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getEquipos = new List<string>();

            while (odr.Read())
            {
                getEquipos.Add("" + odr["equip_no"]);
            }
            return getEquipos;
        }

        public List<string> GetInspectores()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var sqlQuery = "SELECT TRIM(TABLE_CODE)||' - '||TABLE_DESC AS INSP FROM ellipse.msf010 WHERE table_type='VI' AND SUBSTR(TABLE_DESC,1,6)<='999999' ORDER BY TABLE_CODE";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getInspectores = new List<string>();
           
            while (odr.Read())
            {
                getInspectores.Add("" + odr["INSP"]);
            }
            return getInspectores;
        }

        public List<string> GetMonitoreos()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'OI'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getMonitoreos = new List<string>();

            while (odr.Read())
            {
                getMonitoreos.Add("" + odr["table_code"]);
            }
            return getMonitoreos;
        }

        public List<string> GetMonitoreosMtto()
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'OI' and trim(table_code) in ('IE','UT')";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getMonitoreos = new List<string>();

            while (odr.Read())
            {
                getMonitoreos.Add("" + odr["table_code"]);
            }
            return getMonitoreos;
        }

        public void changesGroupRange_Change(Excel.Range target)
        {
            var currentRow = 14;
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                string sqlQuery = "select " +
                                     "trim(substr(m.comp_mod_data,1,4)||' - '||(select table_desc from ellipse.msf010 where table_type = 'CO' and trim(table_code) = trim(substr(m.comp_mod_data,1,4)))) as comp, " +
                                     "trim(substr(m.comp_mod_data,5,2)||' - '||(select table_desc from ellipse.msf010 where table_type = 'MO' and trim(table_code) = trim(substr(m.comp_mod_data,5,2)))) as mod, " +
                                     "trim(trim(m.cond_mon_pos)||' - '||(select table_desc from ellipse.msf010 where table_type = 'PM' and trim(table_code) = trim(m.cond_mon_pos))) as pos, " +
                                     "trim(m.cond_mon_meas) as codigo, " +
                                     "trim(t.table_desc) as descripcion, " +
                                     "trim(m.meas_caut_lowr) as caution_low, " +
                                     "trim(m.meas_caut_uppr) as caution, " +
                                     "trim(m.meas_dang_lowr) as danger_low, " +
                                     "trim(m.meas_dang_uppr) as danger " +
                                     "from " +
                                     "ellipse.msf341 m, " +
                                     "ellipse.msf340_set_def s, " +
                                     "ellipse.msf010 t " +
                                     "where " +
                                     "m.cond_mon_type = '" + _cells.GetCell(target.Column, target.Row - 5).Value + "'" +
                                     "and " +
                                     "( " +
                                     "m.type_reference = 'G'||(select equip_grp_id from ellipse.msf600 where equip_no = '" + target.Value + "')" +
                                     "or m.type_reference = 'E'||'" + target.Value + "'" +
                                     ") " +
                                     "and t.table_type = 'MS' " +
                                     "and t.table_code = m.cond_mon_meas " +
                                     "and s.type_reference = m.type_reference " +
                                     "and s.cond_mon_type = m.cond_mon_type " +
                                     "and s.comp_mod_data = m.comp_mod_data " +
                                     "and s.cond_mon_pos = m.cond_mon_pos " +
                                     "and s.status_340 = 'A' " +
                                     "order by 1,2,3,4";

                var odr = _eFunctions.GetQueryResult(sqlQuery);

                Limpiar();

                while (odr.Read())
                {
                        _cells.GetCell("A" + currentRow).Value = odr["comp"] + "";
                        _cells.GetCell("B" + currentRow).Value = odr["mod"] + "";
                        _cells.GetCell("C" + currentRow).Value = odr["pos"] + "";
                        _cells.GetCell("D" + currentRow).Value = odr["codigo"] + "";
                        _cells.GetCell("E" + currentRow).Value = odr["descripcion"] + "";
                        _cells.GetCell("F" + currentRow).Value = odr["caution_low"] + "";
                        _cells.GetCell("G" + currentRow).Value = odr["caution"] + "";
                        _cells.GetCell("H" + currentRow).Value = odr["danger_low"] + "";
                        _cells.GetCell("I" + currentRow).Value = odr["danger"] + "";

                        currentRow++;                        
                        _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();                    
                }

                Centrar();

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        public void changesGroupRange_Change2(Excel.Range target)
        {
            var currentRow = 14;
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                string sqlQuery = "select " +
                                     "trim(substr(m.comp_mod_data,1,4)||' - '||(select table_desc from ellipse.msf010 where table_type = 'CO' and trim(table_code) = trim(substr(m.comp_mod_data,1,4)))) as comp, " +
                                     "trim(substr(m.comp_mod_data,5,2)||' - '||(select table_desc from ellipse.msf010 where table_type = 'MO' and trim(table_code) = trim(substr(m.comp_mod_data,5,2)))) as mod, " +
                                     "trim(trim(m.cond_mon_pos)||' - '||(select table_desc from ellipse.msf010 where table_type = 'PM' and trim(table_code) = trim(m.cond_mon_pos))) as pos, " +
                                     "trim(m.cond_mon_meas) as codigo, " +
                                     "trim(t.table_desc) as descripcion, " +
                                     "trim(m.meas_caut_lowr) as caution_low, " +
                                     "trim(m.meas_caut_uppr) as caution, " +
                                     "trim(m.meas_dang_lowr) as danger_low, " +
                                     "trim(m.meas_dang_uppr) as danger " +
                                     "from " +
                                     "ellipse.msf341 m, " +
                                     "ellipse.msf340_set_def s, " +
                                     "ellipse.msf010 t " +
                                     "where " +
                                     "m.cond_mon_type = '" + target.Value + "'" +
                                     "and " +
                                     "( " +
                                     "m.type_reference = 'G'||(select equip_grp_id from ellipse.msf600 where equip_no = '" + _cells.GetCell(target.Column, target.Row + 5).Value + "')" +
                                     "or m.type_reference = 'E'||'" + _cells.GetCell(target.Column, target.Row + 5).Value + "'" +
                                     ") " +
                                     "and t.table_type = 'MS' " +
                                      "and t.table_code = m.cond_mon_meas " +
                                     "and s.type_reference = m.type_reference " +
                                     "and s.cond_mon_type = m.cond_mon_type " +
                                     "and s.comp_mod_data = m.comp_mod_data " +
                                     "and s.cond_mon_pos = m.cond_mon_pos " +
                                     "and s.status_340 = 'A' " +
                                     "order by 1,2,3,4";

                var odr = _eFunctions.GetQueryResult(sqlQuery);

                Limpiar();

                while (odr.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = odr["comp"] + "";
                    _cells.GetCell("B" + currentRow).Value = odr["mod"] + "";
                    _cells.GetCell("C" + currentRow).Value = odr["pos"] + "";
                    _cells.GetCell("D" + currentRow).Value = odr["codigo"] + "";
                    _cells.GetCell("E" + currentRow).Value = odr["descripcion"] + "";
                    _cells.GetCell("F" + currentRow).Value = odr["caution_low"] + "";
                    _cells.GetCell("G" + currentRow).Value = odr["caution"] + "";
                    _cells.GetCell("H" + currentRow).Value = odr["danger_low"] + "";
                    _cells.GetCell("I" + currentRow).Value = odr["danger"] + "";

                    currentRow++;
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                Centrar();

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        public void changesGroupRange_ChangeMTTO(Excel.Range target)
        {
            var currentRow = 15;
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                string sqlQuery = "select " +                                     
                                     "trim(m.cond_mon_meas) as codigo, " +
                                     "trim(t.table_desc) as descripcion, " +
                                     "trim(m.meas_caut_lowr) as caution_low, " +
                                     "trim(m.meas_caut_uppr) as caution, " +
                                     "trim(m.meas_dang_lowr) as danger_low, " +
                                     "trim(m.meas_dang_uppr) as danger " +
                                     "from " +
                                     "ellipse.msf341 m, " +
                                     "ellipse.msf340_set_def s, " +
                                     "ellipse.msf010 t " +
                                     "where " +
                                     "m.cond_mon_type = '" + _cells.GetCell(target.Column, target.Row - 1).Value + "'" +
                                     "and " +
                                     "( " +
                                     "m.type_reference = 'G'||(select equip_grp_id from ellipse.msf600 where equip_no = '" + target.Value + "')" +
                                     "or m.type_reference = 'E'||'" + target.Value + "'" +
                                     ") " +
                                     "and t.table_type = 'MS' " +
                                     "and t.table_code = m.cond_mon_meas " +
                                     "and s.type_reference = m.type_reference " +
                                     "and s.cond_mon_type = m.cond_mon_type " +
                                     "and s.comp_mod_data = m.comp_mod_data " +
                                     "and s.cond_mon_pos = m.cond_mon_pos " +
                                     "and s.status_340 = 'A' " +
                                     "order by 1,2,3,4";

                var odr = _eFunctions.GetQueryResult(sqlQuery);

                LimpiarMtto();

                while (odr.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = odr["codigo"] + "";
                    _cells.GetCell("B" + currentRow).Value = odr["descripcion"] + "";
                    _cells.GetCell("C" + currentRow).Value = odr["caution_low"] + "";
                    _cells.GetCell("D" + currentRow).Value = odr["caution"] + "";
                    _cells.GetCell("E" + currentRow).Value = odr["danger_low"] + "";
                    _cells.GetCell("F" + currentRow).Value = odr["danger"] + "";

                    currentRow++;
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                CentrarMtto();
                ValidacionMtto();

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        public void changesGroupRange_Change2MTTO(Excel.Range target)
        {
            var currentRow = 15;
            try
            {
                _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);

                string sqlQuery = "select " +
                                     "trim(m.cond_mon_meas) as codigo, " +
                                     "trim(t.table_desc) as descripcion, " +
                                     "trim(m.meas_caut_lowr) as caution_low, " +
                                     "trim(m.meas_caut_uppr) as caution, " +
                                     "trim(m.meas_dang_lowr) as danger_low, " +
                                     "trim(m.meas_dang_uppr) as danger " +
                                     "from " +
                                     "ellipse.msf341 m, " +
                                     "ellipse.msf340_set_def s, " +
                                     "ellipse.msf010 t " +
                                     "where " +
                                     "m.cond_mon_type = '" + target.Value + "'" +
                                     "and " +
                                     "( " +
                                     "m.type_reference = 'G'||(select equip_grp_id from ellipse.msf600 where equip_no = '" + _cells.GetCell(target.Column, target.Row + 1).Value + "')" +
                                     "or m.type_reference = 'E'||'" + _cells.GetCell(target.Column, target.Row + 1).Value + "'" +
                                     ") " +
                                     "and t.table_type = 'MS' " +
                                      "and t.table_code = m.cond_mon_meas " +
                                     "and s.type_reference = m.type_reference " +
                                     "and s.cond_mon_type = m.cond_mon_type " +
                                     "and s.comp_mod_data = m.comp_mod_data " +
                                     "and s.cond_mon_pos = m.cond_mon_pos " +
                                     "and s.status_340 = 'A' " +
                                     "order by 1,2,3,4";

                var odr = _eFunctions.GetQueryResult(sqlQuery);

                LimpiarMtto();

                while (odr.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = odr["codigo"] + "";
                    _cells.GetCell("B" + currentRow).Value = odr["descripcion"] + "";
                    _cells.GetCell("C" + currentRow).Value = odr["caution_low"] + "";
                    _cells.GetCell("D" + currentRow).Value = odr["caution"] + "";
                    _cells.GetCell("E" + currentRow).Value = odr["danger_low"] + "";
                    _cells.GetCell("F" + currentRow).Value = odr["danger"] + "";

                    currentRow++;
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                CentrarMtto();
                ValidacionMtto();

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A" + currentRow).Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        public void changesGroupRange_FLOTA(Excel.Range target)
        {
            try
            {
                _eFunctions.SetDBSettings(EnviromentConstants.SigcorProductivo);

                string sqlQuery = "SELECT EQU FROM EQMTLIST WHERE FLOTA_ELLIPSE = '" + target.Value + "'" + " AND ACTIVE_FLG = 'Y' ORDER BY 1 ";                    

                var odr = _eFunctions.GetQueryResult(sqlQuery);

 
            var getFlotas = new List<string>();

            while (odr.Read())
            {
                getFlotas.Add("" + odr["EQU"]);
            }

            //Cells.getCell("B12").Value2 = "";
            _cells.SetValidationList(_cells.GetCell("B8"), getFlotas);

            }
            catch (NullReferenceException)
            {
                _cells.GetCell("A15").Value = "No fue Posible Obtener Informacion!";
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
        }

        private void fMantto_Click(object sender, RibbonControlEventArgs e)
        {
            SetSheetHeaderDataMtto();
            Centrar();
            ValidacionMtto();
            _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            var groupCells = _worksheet.Range["B8:B8"];
            var groupRange = _worksheet.Controls.AddNamedRange(groupCells, "GroupRange");

            groupRange.Change += changesGroupRange_ChangeMTTO;

            var groupCells2 = _worksheet.Range["B7:B7"];
            var groupRange2 = _worksheet.Controls.AddNamedRange(groupCells2, "GroupRange2");

            groupRange2.Change += changesGroupRange_Change2MTTO;

            var groupCells3 = _worksheet.Range["B10:B12"];
            var groupRange3 = _worksheet.Controls.AddNamedRange(groupCells3, "GroupRange3");

            groupRange3.Change += AutoAjuste;

            var groupCells4 = _worksheet.Range["B6:B6"];
            var groupRange4 = _worksheet.Controls.AddNamedRange(groupCells4, "GroupRange4");

            groupRange4.Change += changesGroupRange_FLOTA;

        }

        public void SetSheetHeaderDataMtto()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                var excelBook = _excelApp.Workbooks.Add();
                Excel.Worksheet excelSheet = excelBook.ActiveSheet;                
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01Mtto;

                _worksheet = Globals.Factory.GetVstoObject(excelSheet);

                if (_cells == null)

                    _cells = new ExcelStyleCells(_excelApp);

                _cells.GetCell("I1").Value = "OBLIGATORIO";
                _cells.GetCell("I1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("I1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("I2").Value = "OPCIONAL";
                _cells.GetCell("I2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("I3").Value = "INFORMATIVO";
                _cells.GetCell("I3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("I4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("I4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("I5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("I5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange("J1", "XFD1048576").Columns.Hidden = true;
                                                
                _cells.GetCell("A14").Value = "CODIGO";
                _cells.GetCell("A14").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("B14").Value = "DESCRIPCION";
                _cells.GetCell("B14").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("C14").Value = "CAUTION LOW";
                _cells.GetCell("C14").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("D14").Value = "CAUTION";
                _cells.GetCell("D14").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("E14").Value = "DANGER LOW";
                _cells.GetCell("E14").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("F14").Value = "DANGER";
                _cells.GetCell("F14").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                                 
                _cells.GetCell("G14").Value = "VALOR ENCONTRADO";
                _cells.GetCell("G14").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                                 
                _cells.GetCell("H14").Value = "COMENTARIO";
                _cells.GetCell("H14").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                                 
                _cells.GetCell("I14").Value = "RESULTADO";
                _cells.GetCell("I14").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "MONITOREO DE CONDICIONES - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "H5");
                _cells.GetRange("B1", "H5").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("B1", "H5").Borders.Weight = "2";
                _cells.MergeCells("C6", "I12");
                _cells.GetRange("C6", "I12").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("C6", "I12").Borders.Weight = "2";
                _cells.MergeCells("A13", "I13");
                _cells.GetRange("A13", "I13").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetRange("A13", "I13").Borders.Weight = "2";

                _cells.GetCell("A6").Value = "FLOTA";
                _cells.GetCell("A6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationList(_cells.GetCell("B6"), GetFlotas());
                _cells.GetCell("B6").NumberFormat = "@";
                //Cells.setValidationList(Cells.getCell("B11"), getEquipos());
                _cells.GetCell("B6").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B6").Borders.Weight = "2";
                _cells.GetCell("B6").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B6").Font.Bold = true;
                
                _cells.GetCell("A7").Value = "MONITOREO";
                _cells.GetCell("A7").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.SetValidationList(_cells.GetCell("B7"), GetMonitoreosMtto());
                _cells.GetCell("B7").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B7").Borders.Weight = "2";
                _cells.GetCell("B7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B7").Font.Bold = true;

                _cells.GetCell("A8").Value = "EQUIPO";
                _cells.GetCell("A8").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("B8").NumberFormat = "@";
                //Cells.setValidationList(Cells.getCell("B11"), getEquipos());
                _cells.GetCell("B8").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B8").Borders.Weight = "2";
                _cells.GetCell("B8").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B8").Font.Bold = true;

                _cells.GetCell("A9").Value = "FECHA";
                _cells.GetCell("A9").Style = _cells.GetStyle(StyleConstants.TitleRequired);                
                _cells.GetCell("B9").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                _cells.GetCell("B9").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B9").Borders.Weight = "2";
                _cells.GetCell("B9").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                _cells.GetCell("B9").Font.Bold = true;                                

                var inspectores = GetInspectores();

                _cells.GetCell("A10").Value = "INSPECTOR 1";
                _cells.GetCell("A10").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B10"), inspectores);
                _cells.GetCell("B10").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B10").Borders.Weight = "2";
                _cells.GetCell("B10").Font.Bold = true;

                _cells.GetCell("A11").Value = "INSPECTOR 2";
                _cells.GetCell("A11").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B11"), inspectores);
                _cells.GetCell("B11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B11").Borders.Weight = "2";
                _cells.GetCell("B11").Font.Bold = true;

                _cells.GetCell("A12").Value = "INSPECTOR 3";
                _cells.GetCell("A12").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("B12"), inspectores);
                _cells.GetCell("B12").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                _cells.GetCell("B12").Borders.Weight = "2";
                _cells.GetCell("B12").Font.Bold = true;               

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _fechaCalendario = _worksheet.Controls.AddDateTimePicker(_cells.GetCell("B9"), "Calendario");

                _fechaCalendario.ValueChanged += CambioFecha;
                
                _cells.GetCell("B9").Value = _fechaCalendario.Value.ToString("yyyyMMdd"); 

                _cells.GetCell("B9").Select();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void CambioFecha(object sender, EventArgs e)
        {
            var picker = (Microsoft.Office.Tools.Excel.Controls.DateTimePicker)sender;
            _cells.GetCell("B9").Value = picker.Value.ToString("yyyyMMdd"); 


        }

       public void ValidacionMtto()
        {
            _cells.GetCell("G15:G200").Validation.Delete();
            _cells.GetCell("G15:G200").Validation.Add(
                        Excel.XlDVType.xlValidateWholeNumber,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlGreaterEqual,
                        "0");
        }

       private void btnAbout_Click(object sender, RibbonControlEventArgs e)
       {
           new AboutBoxExcelAddIn().ShowDialog();
       }

    }
}
//Desarrollado por Ing. Hussein Villamizar - Septiembre de 2015
//Actualizado por Ing. Héctor Hernández - Mayo de 2017