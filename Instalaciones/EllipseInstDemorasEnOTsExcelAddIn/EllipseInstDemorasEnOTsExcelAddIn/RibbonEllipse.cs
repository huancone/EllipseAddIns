using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using EllipseWorkOrdersClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Worksheet = Microsoft.Office.Tools.Excel.Worksheet;

namespace EllipseInstDemorasEnOTsExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Application _excelApp;

        private const string SheetName01 = "DEMORAS_CARGUE";
        private const string SheetName02 = "DEMORAS_CONSULTA";
        private const string DistrictCode = "INST";
        private const int TitleRow01 = 5;
        private const int TitleRow02 = 5;
        private const int ResultColumn01 = 8;
        private const int ResultColumn02 = 13;
        private const string TableName01 = "DemorasCargueTable";
        private const string TableName02 = "DemorasConsultaTable";
        private const string ValidationSheetName = "ValidationSheet";

        private string _workGroup;

        //List<string> _actionList;
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

        private void btnFormatSheetImis_Click(object sender, RibbonControlEventArgs e)
        {
            _workGroup = DemorasGroups.ImisName;
            FormatSheet();
        }
        private void btnFormatSheetAires_Click(object sender, RibbonControlEventArgs e)
        {
            _workGroup = DemorasGroups.AiresName;
            FormatSheet();
        }

        private void btnConsultarDemoras_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                ConsultarDemoras();
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }

        private void btnUpdateDemoras_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                CreateDuration();
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        private void btnEliminar_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName02)
                DeleteDurations();
            else
                MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
        }
        public void FormatSheet()
        {
            try
            {
                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("A3").Value = _workGroup;
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A3", "G3");
                _cells.GetCell("C1").Value = "CARGUE DE DEMORAS Y ESPERAS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "G2");

                _cells.GetCell("H1").Value = "OBLIGATORIO";
                _cells.GetCell("H1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("H2").Value = "OPCIONAL";
                _cells.GetCell("H2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("H3").Value = "INFORMATIVO";
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                //GENERAL
                _cells.GetCell(1, TitleRow01).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow01).Value = "CONTACTO";
                _cells.GetCell(3, TitleRow01).Value = "EQUIPO";
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPCIÓN";
                _cells.GetCell(5, TitleRow01).Value = "DEMORA (COMPL.CODE)";
                _cells.GetCell(6, TitleRow01).Value = "INICIO";
                _cells.GetCell(7, TitleRow01).Value = "FIN";
                _cells.GetCell(8, TitleRow01).Value = "RESULTADO";
                _cells.GetRange(1, TitleRow01, 7, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.GetRange(2, TitleRow01, 4, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(8, TitleRow01).Style = StyleConstants.TitleResult;


                _cells.SetValidationList(_cells.GetCell(5, TitleRow01 + 1), MyUtilities.GetCodeList(GetCompleteCodeList()), ValidationSheetName, 1);

                Worksheet vstoSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

                var orderNameRange = vstoSheet.Controls.AddNamedRange(_cells.GetRange(1, TitleRow01 + 1, 1, 100), "WorkOrderRange");
                orderNameRange.Change += GetWorkOrderDescriptionChangedValue;

                var demoraDurationRange = vstoSheet.Controls.AddNamedRange(_cells.GetRange(5, TitleRow01 + 1, 5, 100), "demoraDurationRange");
                demoraDurationRange.Change += GetDemoraStartDurationChangedValue;

                _cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                //CONSTRUYO LA HOJA 2
                // ReSharper disable once UseIndexedProperty
                _excelApp.ActiveWorkbook.Sheets.get_Item(2).Activate();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName02;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");
                _cells.GetCell("A3").Value = _workGroup;
                _cells.GetCell("A3").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A3", "J3");
                _cells.GetCell("C1").Value = "CONSULTA DE DEMORAS Y ESPERAS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "INFORMATIVO";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("K3").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleAction);

                _cells.GetCell(1, TitleRow02).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow02).Value = "CONTACTO";
                _cells.GetCell(3, TitleRow02).Value = "EQUIP_NO";
                _cells.GetCell(4, TitleRow02).Value = "DESCRIPCIÓN";
                _cells.GetCell(5, TitleRow02).Value = "DEMORA";
                _cells.GetCell(6, TitleRow02).Value = "INICIO";
                _cells.GetCell(7, TitleRow02).Value = "FIN";
                _cells.GetCell(8, TitleRow02).Value = "DIF.DIAS";
                _cells.GetCell(9, TitleRow02).Value = "ESTADO";
                _cells.GetCell(10, TitleRow02).Value = "TIPO_CIERRE";
                _cells.GetCell(11, TitleRow02).Value = "COMPLETE_CODE";
                _cells.GetCell(12, TitleRow02).Value = "ACCION";
                _cells.GetCell(13, TitleRow02).Value = "RESULTADO";
                _cells.GetRange(1, TitleRow02, 12, TitleRow02).Style = StyleConstants.TitleInformation;
                _cells.GetCell(1, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetRange(5, TitleRow02, 7, TitleRow02).Style = StyleConstants.TitleRequired;
                _cells.GetCell(12, TitleRow02).Style = StyleConstants.TitleAction;
                _cells.GetCell(13, TitleRow02).Style = StyleConstants.TitleResult;
                var actionList = new List<string> { "", "Eliminar" };
                _cells.SetValidationList(_cells.GetCell(12, TitleRow02 + 1), actionList, ValidationSheetName, 2);
                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow02, ResultColumn02, TitleRow02 + 1), TableName02);

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:formatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }
        public void ConsultarDemoras()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _cells.ClearRange("A6", "AZ65536");
                _cells.GetRange(1, TitleRow02 + 1, ResultColumn02, TitleRow02 + 101).NumberFormat =
                    NumberFormatConstants.Text;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var groupName = _cells.GetNullIfTrimmedEmpty(_cells.GetCell("A3").Value);
                var workGroup = "";

                if (groupName != null && groupName.Equals(DemorasGroups.AiresName))
                    workGroup = DemorasGroups.AiresCode;
                if (groupName != null && groupName.Equals(DemorasGroups.ImisName))
                    workGroup = DemorasGroups.ImisCode;

                var sqlQuery = Queries.GetConsultaQuery("INST", workGroup, _eFunctions.DbReference, _eFunctions.DbLink,
                    new List<string>(GetCompleteCodeList().Keys));

                var drDemoras = _eFunctions.GetQueryResult(sqlQuery);

                if (drDemoras != null && !drDemoras.IsClosed)
                {
                    var i = TitleRow02 + 1;
                    while (drDemoras.Read())
                    {
                        _cells.GetCell(1, i).Value = drDemoras["WORK_ORDER"].ToString().Trim();
                        _cells.GetCell(2, i).Value = drDemoras["CONTACTO"].ToString().Trim();
                        _cells.GetCell(3, i).Value = drDemoras["EQUIP_NO"].ToString().Trim();
                        _cells.GetCell(4, i).Value = drDemoras["DESCRIPCION"].ToString().Trim();
                        _cells.GetCell(5, i).Value = drDemoras["DEMORA"].ToString().Trim();
                        _cells.GetCell(6, i).Value = drDemoras["INICIO"].ToString().Trim();
                        _cells.GetCell(7, i).Value = drDemoras["FIN"].ToString().Trim();
                        _cells.GetCell(8, i).Value = drDemoras["DIFERENCIA_DIAS"].ToString().Trim();
                        _cells.GetCell(9, i).Value = drDemoras["ESTADO"].ToString().Trim();
                        _cells.GetCell(10, i).Value = drDemoras["TIPO_CIERRE"].ToString().Trim();
                        _cells.GetCell(11, i).Value = drDemoras["COMPLETED_CODE"].ToString().Trim();

                        i++;
                    }
                }
                else
                    MessageBox.Show(@"No se han encontrado datos para el grupo especificado");

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:consultarDemoras()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar obtener datos del grupo seleccionado");
            }
            finally
            {
				_eFunctions.CloseConnection();
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void CreateDuration()
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");
                if (string.IsNullOrWhiteSpace(drpEnvironment.SelectedItem.Label))
                    throw new Exception("\nSeleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);


                var i = TitleRow01 + 1;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
                var today =
                    Convert.ToDouble(string.Format("{0:0000}", DateTime.Now.Year) +
                                     string.Format("{0:00}", DateTime.Now.Month) +
                                     string.Format("{0:00}", DateTime.Now.Day));

                while ("" + _cells.GetCell(1, i).Value != "")
                {
                    try
                    {
                        string workOrder = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                        string durationCode = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value); //demora
                        string startDate = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value);
                        string endDate = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value);

                        if (string.IsNullOrWhiteSpace(durationCode) || durationCode.Length < 3)
                            throw new NullReferenceException("Código de demora no es válido");
                        durationCode = durationCode.Substring(0, 3); //para obtener solo el código

                        var startDuration = new WorkOrderDuration
                        {
                            jobDurationsCode = durationCode,
                            jobDurationsDate = startDate,
                            jobDurationsStart = "000000",
                            jobDurationsFinish = "000000"
                        };

                        var endDuration = new WorkOrderDuration
                        {
                            jobDurationsCode = durationCode,
                            jobDurationsDate = endDate,
                            jobDurationsStart = "000000",
                            jobDurationsFinish = "010000"
                        };

                        //se valida este campo para algunas comparaciones numéricas, pero en el elemento endDuration va como original
                        if (string.IsNullOrWhiteSpace(endDate))
                            endDate = "000000";

                        //la fecha de inicio no puede ser nula
                        if (string.IsNullOrWhiteSpace(startDate))
                            throw new NullReferenceException(
                                "La fecha de inicio no puede ser nula. Tiene que ser una fecha válida en formato yyyyMMdd");
                        //las fechas tienen que ser menores o iguales al día de hoy
                        if (Convert.ToDouble(startDate) > today || Convert.ToDouble(endDate) > today)
                            throw new InvalidOperationException(
                                "Fecha de Inicio y Fecha Final no pueden referirse a una fecha futura");

                        //obtengo todas las duraciones de la orden
                        var woDurations = WorkOrderActions.GetWorkOrderDurations(urlService, opSheet, DistrictCode,
                            WorkOrderActions.GetNewWorkOrderDto(workOrder));
                        //filtro solo las que coinciden al código y las pongo en format de DateRange
                        var durationList = GetDurationCodeDateRanges(woDurations.ToList(), durationCode);

                        //valido que la orden no tenga un periodo sin cerrar del mismo código
                        if (durationList.Any() && durationList.ToArray()[durationList.Count() - 1].EndDate == null &&
                            Convert.ToDouble(durationList.ToArray()[durationList.Count() - 1].StartDate) <
                            Convert.ToDouble(startDate))
                            throw new InvalidOperationException(
                                "Existe un periodo sin cerrar de la misma demora con una fecha anterior a la que quiere ingresar");

                        //valido que la fecha de inicio y la fecha final no estén dentro de algún rango ya ingresdo
                        if (!ValidateDateNotInRange(durationList, startDate))
                            throw new InvalidOperationException(
                                "La fecha de inicio no puede estar dentro del intervalo de una demora ya existente con el mismo código");
                        if (!ValidateDateNotInRange(durationList, endDate))
                            throw new InvalidOperationException(
                                "La fecha final no puede estar dentro del intervalo de una demora ya existente con el mismo código");
                        //Si solo estoy ingresando fecha inicial y esta ya existe
                        if (string.IsNullOrWhiteSpace(endDate) || endDate.Equals("000000"))
                        {
                            //fecha inicial está duplicada
                            if (!ValidateDateNotDuplicated(durationList, startDate, 1))
                                throw new InvalidOperationException("Ya existe una demora con esa fecha inicial");
                            WorkOrderActions.CreateWorkOrderDuration(urlService, opSheet, DistrictCode,
                                WorkOrderActions.GetNewWorkOrderDto(workOrder), startDuration);
                        }
                        else //si estoy ingresando ambas fechas
                        {
                            var startIsValid = ValidateDateNotDuplicated(durationList, startDate, 0);
                            var endIsValid = ValidateDateNotDuplicated(durationList, endDate, 1);
                            if (!endIsValid)
                                throw new InvalidOperationException("Ya existe una demora con esa fecha final");
                            if (startIsValid) //creo la fecha inicial
                                WorkOrderActions.CreateWorkOrderDuration(urlService, opSheet, DistrictCode,
                                    WorkOrderActions.GetNewWorkOrderDto(workOrder), startDuration);
                            //creo la fecha final
                            WorkOrderActions.CreateWorkOrderDuration(urlService, opSheet, DistrictCode,
                                WorkOrderActions.GetNewWorkOrderDto(workOrder), endDuration);
                        }
                        _cells.GetCell(ResultColumn01, i).Value = "CARGADO";
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    }
                    i++;
                } //aquí terminaría el while

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:createDuration()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public void DeleteDurations()
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name != SheetName01)
                    throw new Exception("La hoja seleccionada no coincide con el modelo requerido");
                if (string.IsNullOrWhiteSpace(drpEnvironment.SelectedItem.Label))
                    throw new Exception("\nSeleccione un ambiente válido");

                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);


                var i = TitleRow02 + 1;

                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);

                while ("" + _cells.GetCell(1, i).Value != "")
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace("" + _cells.GetCell(12, i).Value))
                        {
                            WorkOrderDTO wo = WorkOrderActions.GetNewWorkOrderDto("" + _cells.GetCell(1, i).Value);
                            var durationStart = new WorkOrderDuration();
                            string durationCode = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);
                            durationCode = durationCode.Substring(0, 3); //para obtener solo el código

                            durationStart.jobDurationsCode = durationCode;
                            durationStart.jobDurationsDate = _cells.GetEmptyIfNull(_cells.GetCell(6, i).Value);
                            durationStart.jobDurationsStart = "000000";
                            durationStart.jobDurationsFinish = "000000";

                            var durationEnd = new WorkOrderDuration
                            {
                                jobDurationsCode = durationCode,
                                jobDurationsDate = _cells.GetEmptyIfNull(_cells.GetCell(7, i).Value),
                                jobDurationsStart = "000000",
                                jobDurationsFinish = "010000"
                            };

                            WorkOrderActions.DeleteWorkOrderDuration(urlService, opSheet, DistrictCode, wo,
                                durationStart);
                            WorkOrderActions.DeleteWorkOrderDuration(urlService, opSheet, DistrictCode, wo, durationEnd);
                            _cells.GetCell(ResultColumn02, i).Value = "ELIMINADO";
                        }

                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(ResultColumn02, i).Value = "ERROR: " + ex.Message;
                    }
                    i++;
                } //aquí terminaría el while

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:DeleteDuration()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
        public List<DateRange> GetDurationCodeDateRanges(List<WorkOrderDuration> durationsWoList, string durationCode)
        {
            //selecciono solo las duraciones que coincidan con el código a buscar
            var durationsCodeList = new List<WorkOrderDuration>();
            foreach (var wod in durationsWoList)
            {
                if (wod.jobDurationsCode == durationCode)
                    durationsCodeList.Add(wod);
            }

            var listRanges = new List<DateRange>();//lista que llevará los rangos
            var newRange = new DateRange(null, null);
            decimal prevValue = 1;//inicia con 1 para garantizar que se inicie siempre con un 0 de apertura
            var i = 0;
            foreach (var wod in durationsCodeList)
            {
                //si es diferente de los valores permitidos para secuencia (0 y 24: punto inicial, 1: punto final
                if (!(wod.jobDurationsHours == 0 || wod.jobDurationsHours == 1 || wod.jobDurationsHours == 24))
                    throw new InvalidOperationException("La Orden contiene un valor de demora no válido");
                //si existen duraciones que no coincidan con el orden establecido
                if (wod.jobDurationsHours == prevValue)
                    throw new InvalidOperationException("La Orden contiene una o más secuencias de demoras inválidas");

                if (i % 2 == 0)//franja inicial
                    newRange = new DateRange(wod.jobDurationsDate, null);
                else//franja final
                {
                    newRange.EndDate = wod.jobDurationsDate;
                    listRanges.Add(newRange);
                }

                //actualizo el comparador
                prevValue = wod.jobDurationsHours;
                if (prevValue == 24)//para garantizar compatibilidad con órdenes antes de ellipse 8.4
                    prevValue = 0;
                i++;
            }

            //si la última demora sigue abierta
            if (newRange.StartDate != null && newRange.EndDate == null)
                listRanges.Add(newRange);

            return listRanges;
        }
        /// <summary>
        /// Establece el resultado de búsqueda de la información de una orden después de que esta es escrita
        /// </summary>
        /// <param name="target"></param>
        void GetWorkOrderDescriptionChangedValue(Range target)
        {
            try
            {
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                _cells.GetCell(target.Column + 1, target.Row).Value = "Buscando Orden...";
                var dr = _eFunctions.GetQueryResult(Queries.GetWoDataQuery(DistrictCode, target.Value, _eFunctions.DbReference, _eFunctions.DbLink));

                if (dr != null && !dr.IsClosed)
                {
                    while (dr.Read())
                    {
                        _cells.GetCell(target.Column + 1, target.Row).Value = dr["CONTACTO"].ToString().Trim();
                        _cells.GetCell(target.Column + 2, target.Row).Value = dr["EQUIP_NO"].ToString().Trim();
                        _cells.GetCell(target.Column + 3, target.Row).Value = dr["DESCRIPCION"].ToString().Trim();
                    }
                }
                else
                    _cells.GetCell(target.Column + 1, target.Row).Value = "Orden no encontrada";

            }
            catch (NullReferenceException)
            {
                _cells.GetCell(target.Column + 1, target.Row).Value = "No fue Posible Obtener Informacion!";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        /// <summary>
        /// Establece el resultado de búsqueda de la información de una demora de una orden después de que esta es escrita
        /// </summary>
        /// <param name="target"></param>
        void GetDemoraStartDurationChangedValue(Range target)
        {
            try
            {
                string workOrder = _cells.GetEmptyIfNull(_cells.GetCell(1, target.Row).Value);
                if (workOrder.Equals("")) return;
                string durationCode = _cells.GetEmptyIfNull(target.Value);
                durationCode = durationCode.Substring(0, 3);
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var dr = _eFunctions.GetQueryResult(Queries.GetWoDurDatesDataQuery(_eFunctions.DbReference, _eFunctions.DbLink, workOrder, durationCode));

                if (dr == null || dr.IsClosed) return;
                dr.Read();
                var endDate = dr["FIN"].ToString().Trim();
                if (endDate.Equals(""))
                    _cells.GetCell(target.Column + 1, target.Row).Value = dr["INICIO"].ToString().Trim();
            }
            catch (Exception error)
            {
                if (Debugger.DebugErrors)
                    MessageBox.Show(error.Message);
            }
        }

        /// <summary>
        /// Valida que la fecha a ingresar no esté dentro de un rango ya existente. Principio de frontera excluyente
        /// </summary>
        /// <param name="listRange">List(DateRange): Listado de elementos (startDate, endDate) a comparar</param>
        /// <param name="dateToValidate">string: fecha en format yyyyMMdd para comparar</param>
        /// <param name="allowLast"></param>
        /// <returns></returns>
        public bool ValidateDateNotInRange(List<DateRange> listRange, string dateToValidate, bool allowLast = false)
        {

            //una fecha final puede ser nula
            if (dateToValidate == null || !listRange.Any())
                return true;

            //una condición especial de frontera final para startDate
            var last = listRange.Last();
            if (allowLast)
                if (last.StartDate == dateToValidate)
                    return true;

            foreach (var range in listRange)
            {
                //si no hay con qué comparar se asume que es válida
                if (range.StartDate == null || range.EndDate == null)
                    return true;

                if (Convert.ToDouble(dateToValidate) >= Convert.ToDouble(range.StartDate) && Convert.ToDouble(dateToValidate) <= Convert.ToDouble(range.EndDate))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listRange"></param>
        /// <param name="dateToValidate"></param>
        /// <param name="value">int: 1 si es punto inicial, 0 y 24 si es punto final</param>
        /// <returns>bool: true si no está duplicado, false si está duplicado</returns>
        public bool ValidateDateNotDuplicated(List<DateRange> listRange, string dateToValidate, int value)
        {
            if (!listRange.Any())
                return true;
            //si no coincide con el código
            if (!(value == 0 || value == 1 || value == 24))
                return false;

            foreach (var range in listRange)
            {
                switch (value)
                {
                    case 0:
                    case 24:
                    {
                        if (range.StartDate == dateToValidate)
                            return false;
                        break;
                    }
                    case 1 when range.EndDate == dateToValidate:
                        return false;
                }
            }
            return true;
        }
        public class DateRange
        {
            public string StartDate;
            public string EndDate;

            public DateRange(string startDate, string endDate)
            {
                StartDate = startDate;
                EndDate = endDate;
            }
        }
        public Dictionary<string, string> GetCompleteCodeList()
        {
            var listCode = new Dictionary<string, string>
            {
                {"C01", "Falta disp. sitio/cliente"},
                {"C02", "Pend. entrega de alcance"},
                {"C03", "Alcance ampliado"},
                {"C04", "Proyecto Mayor Alcance"},
                {"C05", "Demora autoriza. paso a paso"},
                {"C06", "Material dificil de conseguir"},
                {"C07", "Pend. intervencion terceros"},
                {"C08", "Pend. permiso o libranza"},
                {"C09", "Autorizado en el fin de semana"},
                {"C10", "Planos iniciales no disponible"},
                {"C11", "Factor Clima"},
                {"C12", "Restricción en tiempos/ejecu"},
                {"C13", "Nodisponib/equiposdeCERREJON"},
                {"C14", "Aprobación/APU(Mater/Rubros)"},
                {"C15", "Demora*aprob/compra*costoreem"},
                {"C16", "Pendiente*autorización/PPTO"},
                {"C17", "Demora*entrega/mat.sumint*CER"},
                {"C18", "O.S. congelada"},
                {"C19", "Programada por el Cliente"},
                {"K01", "Sin recursos"},
                {"K02", "Sin materiales"},
                {"K03", "Sin personal"},
                {"K04", "Sin equipos"},
                {"K05", "Falencias en la programacion"},
                {"K06", "Requisicion no oportuna Mat."},
                {"K07", "Demora elaboración paso a paso"},
                {"K08", "Demoraenproc/compra*costreembo"},
                {"K09", "Pendiente elaboración/ppto."}
            };


            return listCode;
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }




    }

    public static class DemorasGroups
    {
        public static string ImisCode = "CALLCEN";
        public static string ImisName = "IMIS";
        public static string AiresCode = "MANTENIMIENTO DE AIRES";
        public static string AiresName = "AAPREV";

    }

    [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
    internal static class Queries
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="districtCode"></param>
        /// <param name="workGroup"></param>
        /// <param name="dbReference"></param>
        /// <param name="dbLink"></param>
        /// <param name="codeList">Lista de códigos válidos de 3 caracteres</param>
        /// <returns></returns>
        public static string GetConsultaQuery(string districtCode, string workGroup, string dbReference, string dbLink, List<string> codeList)
        {
            var jobDurCodeRestriction = "";
            if (codeList != null && codeList.Any())
            {
                jobDurCodeRestriction = "AND DUR.JOB_DUR_CODE IN (";
                foreach (var code in codeList)
                    jobDurCodeRestriction = jobDurCodeRestriction + " '" + code.Substring(0, 3) + "',";
                jobDurCodeRestriction = jobDurCodeRestriction.Substring(0, jobDurCodeRestriction.Length - 1) + ")";

            }

            var sqlQuery = "" +
                " WITH LDEM AS (" +
                "   SELECT DISTINCT" +
                "     WO.WO_DESC CONTACTO, TRIM (REPLACE (TRIM (COM.STD_VOLAT_1) || ' ' || TRIM (COM.STD_VOLAT_2) || ' ' || TRIM (COM.STD_VOLAT_3) || ' ' || TRIM (COM.STD_VOLAT_4) || ' ' || TRIM (COM.STD_VOLAT_5), '.HEADING', '')) DESCRIPCION," +
                "     DUR.JOB_DUR_DATE, DUR.JOB_DUR_CODE, DECODE(DUR.JOB_DUR_HOURS, 24, 0, DUR.JOB_DUR_HOURS) JOB_DUR_HOURS," +
                "     LEAD (DUR.JOB_DUR_DATE) OVER (PARTITION BY DUR.DSTRCT_CODE, WO.WORK_ORDER, DUR.JOB_DUR_CODE ORDER BY DUR.JOB_DUR_DATE) LEAD_DATE," +
                "     LEAD (DECODE(DUR.JOB_DUR_HOURS, 24, 0, DUR.JOB_DUR_HOURS)) OVER (PARTITION BY DUR.DSTRCT_CODE, WO.WORK_ORDER, DUR.JOB_DUR_CODE ORDER BY DUR.JOB_DUR_DATE) LEAD_HOURS," +
                "     CODE.TABLE_DESC, WO.WORK_ORDER, DUR.WORK_ORDER AS WORK_ORDER1," +
                "     WO.RAISED_DATE AS APERTURA, WO.AUTHSD_DATE, WO.WO_TYPE, WO.WO_TYPE AS WO_TYPE1," +
                "     WO.EQUIP_NO, WO.FINAL_COSTS, WO.WO_STATUS_M, WO.COMPLETED_CODE" +
                "   FROM" +
                "     " + dbReference + ".MSF620" + dbLink + " WO LEFT OUTER JOIN " + dbReference + ".MSF096_STD_VOLAT COM ON COM.STD_KEY = WO.DSTRCT_WO" +
                "     LEFT JOIN " + dbReference + ".MSF622" + dbLink + " DUR ON WO.DSTRCT_CODE  = DUR.DSTRCT_CODE AND DUR.WORK_ORDER = WO.WORK_ORDER" +
                "     LEFT JOIN " + dbReference + ".MSF010" + dbLink + " CODE ON DUR.JOB_DUR_CODE     = CODE.TABLE_CODE" +
                "   WHERE" +
                "     WO.DSTRCT_CODE      = '" + districtCode + "'" +
                "     " + jobDurCodeRestriction + "" +
                "     AND COM.STD_TEXT_CODE   = 'WO'" +
                "     AND WO.WORK_GROUP       = '" + workGroup + "'" +
                "     AND CODE.TABLE_TYPE     = 'JI'" +
                "     AND COM.STD_LINE_NO     = '0000'" +
                "   )" +
                " SELECT LDEM.WORK_ORDER, LDEM.CONTACTO, LDEM.EQUIP_NO, LDEM.DESCRIPCION," +
                "   TRIM (LDEM.JOB_DUR_CODE) || ' - ' || TRIM (LDEM.TABLE_DESC) DEMORA, LDEM.JOB_DUR_DATE INICIO, LDEM.LEAD_DATE FIN," +
                "   TRUNC (TO_DATE (LDEM.LEAD_DATE, 'YYYYMMDD')) - TRUNC (TO_DATE (LDEM.JOB_DUR_DATE, 'YYYYMMDD')) AS DIFERENCIA_DIAS," +
                "   LDEM.WO_STATUS_M AS ESTADO, LDEM.COMPLETED_CODE," +
                "   DECODE (TRIM (LDEM.COMPLETED_CODE), '06', 'CERRADA NORMAL', '08', 'CANCELADA O ANULADA')       AS TIPO_CIERRE" +
                " FROM" +
                "   LDEM" +
                " WHERE" +
                "   LDEM.COMPLETED_CODE NOT IN ('08', 'CN')" +
                "   AND LDEM.JOB_DUR_HOURS = 0 AND TRIM (LDEM.FINAL_COSTS) IS NULL" +
                " ORDER BY LDEM.WORK_ORDER";

            return sqlQuery;
        }
        /// <summary>
        /// Obtiene la información de una WorkOrder especificada
        /// </summary>
        /// <param name="districtCode"></param>
        /// <param name="workOrder"></param>
        /// <param name="dbReference"></param>
        /// <param name="dbLink"></param>
        /// <returns></returns>
        public static string GetWoDataQuery(string districtCode, string workOrder, string dbReference, string dbLink)
        {
            var sqlQuery = "" +
            " SELECT" +
            "   WO.WO_DESC CONTACTO, WO.EQUIP_NO," +
            "   TRIM( REPLACE( TRIM( COM.STD_VOLAT_1 ) || ' ' || TRIM( COM.STD_VOLAT_2 ) || ' ' || TRIM( COM.STD_VOLAT_3 ) || ' ' || TRIM( COM.STD_VOLAT_4 ) || ' ' || TRIM( COM.STD_VOLAT_5 ), '.HEADING', '' ) ) DESCRIPCION" +
            " FROM" +
            "   " + dbReference + ".MSF620" + dbLink + " WO" +
            " LEFT OUTER JOIN " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " COM" +
            " ON" +
            "   COM.STD_KEY = WO.DSTRCT_WO" +
            " WHERE" +
            "   WO.DSTRCT_CODE        = '" + districtCode + "'" +
            "   AND COM.STD_TEXT_CODE = 'WO'" +
            "   AND WO.WORK_ORDER     = '" + workOrder + "'";

            return sqlQuery;
        }

        /// <summary>
        /// Obtiene la información de el último periodo de fechas para una Orden y Código especificado
        /// </summary>
        /// <param name="dbReference"></param>
        /// <param name="dbLink"></param>
        /// <param name="workOrder"></param>
        /// <param name="durationCode">string: Código de tres caracteres de la duración</param>
        /// <returns></returns>
        public static string GetWoDurDatesDataQuery(string dbReference, string dbLink, string workOrder, string durationCode)
        {
            var sqlQuery = "" +
            " WITH LDEM AS" +
            "   (SELECT DUR.JOB_DUR_DATE," +
            "     DUR.JOB_DUR_CODE," +
            "     DECODE(DUR.JOB_DUR_HOURS, 24, 0, DUR.JOB_DUR_HOURS) JOB_DUR_HOURS," +
            "     LEAD (DUR.JOB_DUR_DATE) OVER (PARTITION BY DUR.DSTRCT_CODE, DUR.WORK_ORDER, DUR.JOB_DUR_CODE ORDER BY DUR.JOB_DUR_DATE) LEAD_DATE," +
            "     LEAD (DECODE(DUR.JOB_DUR_HOURS, 24, 0, DUR.JOB_DUR_HOURS)) OVER (PARTITION BY DUR.DSTRCT_CODE, DUR.WORK_ORDER, DUR.JOB_DUR_CODE ORDER BY DUR.JOB_DUR_DATE) LEAD_HOURS," +
            "     DUR.WORK_ORDER," +
            "     DUR.WORK_ORDER AS WORK_ORDER1" +
            "   FROM ELLIPSE.MSF622 DUR" +
            "   WHERE DUR.JOB_DUR_CODE = '" + durationCode + "'" +
            "   AND DUR.WORK_ORDER     = '" + workOrder + "'" +
            "   ORDER BY DUR.JOB_DUR_DATE DESC" +
            "   )" +
            " SELECT LDEM.WORK_ORDER," +
            "   LDEM.JOB_DUR_DATE INICIO," +
            "   LDEM.LEAD_DATE FIN" +
            " FROM LDEM" +
            " WHERE LDEM.JOB_DUR_HOURS = 0" +
            " AND ROWNUM = 1";

            return sqlQuery;
        }
    }
}
