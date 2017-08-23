using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Web.Services.Ellipse;
using Microsoft.Office.Tools.Ribbon;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using EllipseCommonsClassLibrary;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EllipseEqOperStatisticsExcelAddIn.EquipmentOperatingStatisticsService;
using Microsoft.Office.Tools.Excel;

namespace EllipseEqOperStatisticsExcelAddIn
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Local")]
    [SuppressMessage("ReSharper", "UseObjectOrCollectionInitializer")]
    [SuppressMessage("ReSharper", "LoopCanBeConvertedToQuery")]
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Excel.Application _excelApp;

        private const string SheetName01 = "OperationStatistics";
        private const int TitleRow01 = 4;
        private const int ResultColumn01 = 10;
        private const string TableName01 = "OperationStatisticsTable";

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            //Office 2013 requiere no ejecutar esta sentencia al iniciar porque no se cuenta con un libro activo vacío. Se debe ejecutar obligatoriamente al formatear las hojas
            //adcionalmente validar la cantidad de hojas a utilizar al momento de dar formato
            //if (_cells == null)
            //    _cells = new ExcelStyleCells(_excelApp);
            var enviroments = EnviromentConstants.GetEnviromentList();
            foreach (var env in enviroments)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = env;
                drpEnviroment.Items.Add(item);
            }
        }

        private void btnFormatSheet_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheetHeaderData();
        }
        private void btnLoadStatistics_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
            {
                LoadStatistics();
            }
            else
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para el cargue de estadísticas");
        }

        /// <summary>
        /// Establece el formato de la hoja para el cargue de estadísticas de operación
        /// </summary>
        public void FormatSheetHeaderData()
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

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A2");

                _cells.GetCell("B1").Value = "EQUIPMENT OPERATION STATISTICS - ELLIPSE 8";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, TitleRow01, 7, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(1, TitleRow01).Value = "FECHA";
                _cells.GetCell(1, TitleRow01).AddComment("yyyyMMdd");
                _cells.GetCell(2, TitleRow01).Value = "TURNO";
                _cells.GetCell(2, TitleRow01).Style = StyleConstants.TitleOptional;
                _cells.GetCell(3, TitleRow01).Value = "EQUIPO";
                _cells.GetCell(4, TitleRow01).Value = "DESCRIPCIÓN";
                _cells.GetCell(4, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(5, TitleRow01).Value = "TIPO ESTAD.";
                _cells.GetCell(5, TitleRow01).AddComment("Ej: HR");
                _cells.GetCell(6, TitleRow01).Value = "TIPO ENTRADA";
                _cells.GetCell(6, TitleRow01).AddComment("Ej: D - DAYLY, M - METER. Predeterminado M");
                _cells.GetCell(7, TitleRow01).Value = "FECHA ÚLTIMA EST.";
                _cells.GetCell(7, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetCell(8, TitleRow01).Value = "MEDIDOR ÚLT. EST.";
                _cells.GetCell(8, TitleRow01).Style = StyleConstants.TitleInformation;
                _cells.GetRange(1, TitleRow01 + 1, 8, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;

                _cells.GetCell(9, TitleRow01).Value = "VALOR ESTAD.";

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = _cells.GetStyle(StyleConstants.TitleResult);




                var table = _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                //búsquedas especiales de tabla
                var tableObject = Globals.Factory.GetVstoObject(table);
                tableObject.Change += GetTableChangedValue;
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        /// <summary>
        /// Carga la estadística de la hoja especificada
        /// </summary>
        public void LoadStatistics()
        {
            try
            {
                if (drpEnviroment.SelectedItem.Label != null && !drpEnviroment.SelectedItem.Label.Equals(""))
                {
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    _cells.SetCursorWait();
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    var proxySheet = new EquipmentOperatingStatisticsService.EquipmentOperatingStatisticsService();
                    proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) +
                                     "/EquipmentOperatingStatistics";

                    var opSheet = new OperationContext();

                    opSheet.district = _frmAuth.EllipseDsct;
                    opSheet.position = _frmAuth.EllipsePost;
                    opSheet.maxInstances = 100;
                    opSheet.returnWarnings = Debugger.DebugWarnings;
                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
                    const int startRow = 5;
                    var i = startRow;
                    while ("" + _cells.GetCell(1, i).Value != "")
                    {
                        try
                        {

                            var request = new List<EquipmentOperatingStatisticsDTO>();

                            var reqItem = new EquipmentOperatingStatisticsDTO();


                            reqItem.statisticDate = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                            reqItem.shiftCode = _cells.GetNullIfTrimmedEmpty(_cells.GetCell(2, i).Value);
                            reqItem.equipmentNumber = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                            reqItem.operationStatisticType = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);

                            if (_cells.GetEmptyIfNull(_cells.GetCell(6, i).Value) == "D")
                            {
                                var lastMeter =
                                    GetEquipmentLastMeterValue(_cells.GetEmptyIfNull(_cells.GetCell(3, i).Value2),
                                        _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value2),
                                        _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2));
                                reqItem.meterReading = Convert.ToDecimal(_cells.GetCell(9, i).Value) +
                                                       Convert.ToDecimal(lastMeter);
                            }
                            else
                            {
                                reqItem.meterReading = Convert.ToDecimal(_cells.GetCell(9, i).Value);
                            }

                            reqItem.meterReadingSpecified = true;

                            request.Add(reqItem);
                            var replySheet = proxySheet.multipleAdjust(opSheet, request.ToArray());
                            foreach (var reply in replySheet)
                            {
                                if (reply.errors.Length > 0)
                                {
                                    var errors = "";
                                    foreach (var er in reply.errors)
                                        errors = errors + "/" + er.messageText;
                                    throw new Exception(errors);
                                }
                                _cells.GetCell(ResultColumn01, i).Value = "ENVIADO";
                                _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                            }

                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                            throw;
                        }
                        finally
                        {
                            i++;
                        }
                    } //--while de registros
                } //---if no se está en un ambiente válido
                else
                {
                    MessageBox.Show(@"\nSeleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:LoadStatistics()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        /// <summary>
        /// Establece el resultado de búsqueda de la descripción de un equipo después de que este es escrita
        /// </summary>
        /// <param name="target"></param>
        /// <param name="changedRanges"></param>
        void GetTableChangedValue(Excel.Range target, ListRanges changedRanges)//Excel.Range target)
        {
            switch (target.Column)
            {
                case 3:
                    try
                    {
                        _cells.GetCell(target.Column + 1, target.Row).Value = "Buscando Equipo...";
                        string description = GetEquipmentDescription("" + target.Value);

                        _cells.GetCell(target.Column + 1, target.Row).Value = !string.IsNullOrWhiteSpace(description) ? description.Trim() : "Equipo no encontrado";

                        _cells.GetCell(target.Column + 1, target.Row).Columns.AutoFit();

                    }
                    catch (NullReferenceException)
                    {
                        _cells.GetCell(target.Column, target.Row + 1).Value = "No fue Posible Obtener Informacion!";
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.Message);
                    }
                    break;
                case 5:
                    try
                    {
                        var description = GetEquipmentLastStat(_cells.GetCell(3, target.Row).Value2, _cells.GetCell(5, target.Row).Value2, _cells.GetEmptyIfNull(_cells.GetCell(1, target.Row).Value2));

                        _cells.GetCell(7, target.Row).Value = !string.IsNullOrWhiteSpace(description[0]) ? description[0].Trim() : "";
                        _cells.GetCell(7, target.Row).Columns.AutoFit();
                        _cells.GetCell(8, target.Row).Value = !string.IsNullOrWhiteSpace(description[1]) ? description[1].Trim() : "";
                        _cells.GetCell(8, target.Row).Columns.AutoFit();
                    }
                    catch (NullReferenceException)
                    {
                        _cells.GetCell(target.Column, target.Row + 1).Value = "";
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.Message);
                    }
                    break;
            }
        }

        /// <summary>
        /// Obtiene la descripción del equipo a partir del número de equipo
        /// </summary>
        /// <param name="equipNo">string: EquipmentNo para obtener la descripción</param>
        /// <returns>string: EquipmentNo. Null si el equipo no existe</returns>
        string GetEquipmentDescription(string equipNo)
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var dbReference = _eFunctions.dbReference;
            var dbLink = _eFunctions.dbLink;

            var query = "SELECT EQ.* FROM " + dbReference + ".MSF600" + dbLink + " EQ WHERE TRIM(EQ.EQUIP_NO) = '" + equipNo + "'";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drEquipments = _eFunctions.GetQueryResult(query);

            if (drEquipments == null || drEquipments.IsClosed || !drEquipments.HasRows) return null;

            while (drEquipments.Read())
                return ("" + drEquipments["ITEM_NAME_1"]).Trim() + " " + ("" + drEquipments["ITEM_NAME_2"]).Trim();
            return null;
        }

        /// <summary>
        /// Obtiene la descripción del equipo a partir del número de equipo
        /// </summary>
        /// <param name="equipNo">object: EquipmentNo para obtener la descripción</param>
        /// <param name="statType">object: Tipo de estadística a obtener</param>
        /// <param name="statDate">Fecha digitada el la columna 1</param>
        /// <returns>string[2]: {fecha, medidor}. Null si no existe</returns>
        string[] GetEquipmentLastStat(object equipNo, object statType, object statDate)
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var dbReference = _eFunctions.dbReference;
            var dbLink = _eFunctions.dbLink;

            var query = "" +
                            " SELECT" +
                            "   STAT_DATE," +
                            "   METER_VALUE" +
                            " FROM" +
                            "   (" +
                            "     SELECT" +
                            "       METER_VALUE," +
                            "       STAT_DATE," +
                            "       MAX(STAT_DATE) OVER(PARTITION BY EQUIP_NO) MAX_FECHA" +
                            "     FROM" +
                            "       " + dbReference + ".MSF400" + dbLink +
                            "     WHERE" +
                            "       STAT_TYPE = '" + statType + "'" +
                            "     AND KEY_400_TYPE = 'E'" +
                            "     AND EQUIP_NO = '" + equipNo + "' AND STAT_DATE <= '" + statDate + "'" +
                            "   )" +
                            " WHERE" +
                            "   STAT_DATE = MAX_FECHA";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drEquipments = _eFunctions.GetQueryResult(query);

            if (drEquipments == null || drEquipments.IsClosed || !drEquipments.HasRows) return null;

            drEquipments.Read();

            var stat = new string[2];
            stat[0] = ("" + drEquipments["STAT_DATE"]).Trim();
            stat[1] = ("" + drEquipments["METER_VALUE"]).Trim();
            return stat;
        }

        string GetEquipmentLastMeterValue(object equipNo, object statType, object statDate)
        {
            _eFunctions.SetDBSettings(drpEnviroment.SelectedItem.Label);
            var dbReference = _eFunctions.dbReference;
            var dbLink = _eFunctions.dbLink;

            var query = "" +
                            " SELECT" +
                            "   METER_VALUE" +
                            " FROM" +
                            "   (" +
                            "     SELECT" +
                            "       METER_VALUE," +
                            "       STAT_DATE," +
                            "       MAX(STAT_DATE) OVER(PARTITION BY EQUIP_NO) MAX_FECHA" +
                            "     FROM" +
                            "       " + dbReference + ".MSF400" + dbLink +
                            "     WHERE" +
                            "       STAT_TYPE = '" + statType + "'" +
                            "     AND KEY_400_TYPE = 'E'" +
                            "     AND EQUIP_NO = '" + equipNo + "' AND STAT_DATE <= '" + statDate + "'" +
                            "   )" +
                            " WHERE" +
                            "   STAT_DATE = MAX_FECHA";

            query = Utils.ReplaceQueryStringRegexWhiteSpaces(query, "WHERE AND", "WHERE ");

            var drEquipments = _eFunctions.GetQueryResult(query);

            if (drEquipments == null || drEquipments.IsClosed || !drEquipments.HasRows) return null;

            drEquipments.Read();


            var stat = drEquipments["METER_VALUE"].ToString();
            return stat;
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnDelete_Click(object sender, RibbonControlEventArgs e)
        {
            DeleteStatistics();
        }

        private void DeleteStatistics()
        {
            try
            {
                if (drpEnviroment.SelectedItem.Label != null && !drpEnviroment.SelectedItem.Label.Equals(""))
                {
                    if (_cells == null)
                        _cells = new ExcelStyleCells(_excelApp);
                    _cells.SetCursorWait();
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnviroment = drpEnviroment.SelectedItem.Label;

                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                    var i = TitleRow01 + 1;
                    while ("" + _cells.GetCell(1, i).Value != "")
                    {
                        try
                        {
                            var opContext = new Screen.OperationContext
                            {
                                district = _frmAuth.EllipseDsct,
                                position = _frmAuth.EllipsePost,
                                maxInstances = 100,
                                maxInstancesSpecified = true,
                                returnWarnings = Debugger.DebugWarnings,
                                returnWarningsSpecified = true
                            };

                            var proxySheet = new Screen.ScreenService();

                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnviroment.SelectedItem.Label) + "/ScreenService";
                            _eFunctions.RevertOperation(opContext, proxySheet);
                            //ejecutamos el programa
                            Screen.ScreenDTO reply = proxySheet.executeScreen(opContext, "MSO400");
                            //Validamos el ingreso
                            if (reply.mapName != "MSM400A") continue;

                            var statisticDate = _cells.GetEmptyIfNull(_cells.GetCell(1, i).Value);
                            var equipmentNumber = _cells.GetEmptyIfNull(_cells.GetCell(3, i).Value);
                            var operationStatisticType = _cells.GetEmptyIfNull(_cells.GetCell(5, i).Value);

                            var arrayFields = new ArrayScreenNameValue();
                            arrayFields.Add("OPTION1I", "3");
                            arrayFields.Add("STAT_DATE1I", statisticDate);
                            arrayFields.Add("STAT_TYPE1I", operationStatisticType);
                            arrayFields.Add("PLANT_NO1I", equipmentNumber);


                            var request = new Screen.ScreenSubmitRequestDTO
                            {
                                screenFields = arrayFields.ToArray(),
                                screenKey = "1"
                            };
                            reply = proxySheet.submit(opContext, request);

                            if (reply != null && !_eFunctions.CheckReplyError(reply) && !_eFunctions.CheckReplyWarning(reply))
                            {
                                arrayFields = new ArrayScreenNameValue();
                                arrayFields.Add("DELETE3I", "Y");

                                request = new Screen.ScreenSubmitRequestDTO
                                {
                                    screenFields = arrayFields.ToArray(),
                                    screenKey = "1"
                                };

                                reply = proxySheet.submit(opContext, request);

                                if (reply != null && (_eFunctions.CheckReplyError(reply) && reply.mapName == "MSM400A"))
                                    throw new ArgumentException(reply.message);

                                _cells.GetCell(ResultColumn01, i).Value = "ELIMINADO";
                                _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                            }
                            else if (reply != null) throw new Exception(reply.message);
                            else throw new Exception(@"No se ha podido obtener respuesta del servidor");
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                            _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                        }
                        finally
                        {
                            i++;
                        }
                    } //--while de registros
                } //---if no se está en un ambiente válido
                else
                {
                    MessageBox.Show(@"\nSeleccione un ambiente válido", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Debugger.LogError("RibbonEllipse:LoadStatistics()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }
    }
}
