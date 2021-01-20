using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using EllipseWorkOrdersClassLibrary;
using EllipseWorkOrdersClassLibrary.WorkOrderService;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace EllipseInstFinalizarInterventoriaExcelAddIn
{
    public partial class RibbonEllipse
    {
        private const string SheetName01 = "FinalizarInterventoria";
        private const int TitleRow01 = 5;
        private const int ResultColumn01 = 12;
        private const string TableName01 = "FinalizarInterTable";
        private const string DistrictCode = "INST";
        private const string WorkGroup = "CALLCEN";
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private Application _excelApp;
        private FormAuthenticate _frmAuth;

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

        private void btnClearSheet_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                CleanSheet();
            else
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para la operación");
        }


        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            if (_excelApp.ActiveWorkbook.ActiveSheet.Name.StartsWith(SheetName01))
                UpdateData();
            else
                MessageBox.Show(@"La hoja de Excel no tiene el formato válido para la operación");
        }

        public void CleanSheet()
        {
            _cells.ClearTableRange(TableName01);
        }

        public void FormatSheet()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;

                //CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "FINALIZAR INTERVENTORIA - MNTTO DE INSTALACIONES";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("K1").Value = "OBLIGATORIO";
                _cells.GetCell("K1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("K2").Value = "OPCIONAL";
                _cells.GetCell("K2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "INFORMATIVO";
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleRequired;

                //GENERAL
                _cells.GetCell(1, TitleRow01).Value = "WORK_ORDER";
                _cells.GetCell(2, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;

                //método antiguo, solo para Interop
                //Worksheet vstoSheet = Globals.Factory.GetVstoObject(_excelApp.ActiveWorkbook.ActiveSheet);
                //vstoSheet.Controls.Remove("SeekOrderData");
                //var orderNameRange =
                //    vstoSheet.Controls.AddNamedRange(Cells.GetRange(1, TitleRow01 + 1, 1, TitleRow01 + 100),
                //        "SeekOrderData");
                //orderNameRange.Change += GetWorkOrderStatusDataChangedValue;

                _cells.GetCell(2, TitleRow01).Value = "FECHA FINALIZACIÓN";
                _cells.GetCell(2, TitleRow01).AddComment("yyyyMMdd");

                _cells.GetCell(3, TitleRow01).Value = "CLIENTE";
                _cells.GetCell(4, TitleRow01).Value = "STATUS";
                _cells.GetCell(5, TitleRow01).Value = "ESTADO";
                _cells.GetCell(6, TitleRow01).Value = "FIRMA CLIENTE";
                _cells.GetCell(7, TitleRow01).Value = "CALIFICACIÓN";
                _cells.GetCell(8, TitleRow01).Value = "NO CONFORMES";
                _cells.GetCell(9, TitleRow01).Value = "VALES ABIERTOS";
                _cells.GetCell(10, TitleRow01).Value = "DEMORAS ABIERTAS";

                _cells.GetRange(3, TitleRow01, 10, TitleRow01).Style = StyleConstants.TitleInformation;

                _cells.GetCell(11, TitleRow01).Value = "MENSAJE FECHA";
                _cells.GetCell(12, TitleRow01).Value = "RESULTADO";
                _cells.GetRange(11, TitleRow01, 12, TitleRow01).Style = StyleConstants.TitleResult;

                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                var table = _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1),
                    TableName01);

                var tableObject = Globals.Factory.GetVstoObject(table);
                tableObject.Change += GetWorkOrderStatusDataChangedValue;
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja");
            }
        }

        public void UpdateData()
        {
            try
            {
                //Se valida la variable de control de acciones en las hojas de excel
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();
                //Se realiza la autenticación
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

                if (_frmAuth.ShowDialog() != DialogResult.OK) return; //no se autentica, se cancela el proceso

                //Instanciar el Contexto de Operación
                var opSheet = new OperationContext
                {
                    district = _frmAuth.EllipseDsct,
                    position = _frmAuth.EllipsePost,
                    maxInstances = 100,
                    maxInstancesSpecified = true,
                    returnWarnings = Debugger.DebugWarnings,
                    returnWarningsSpecified = true
                };

                //Instanciar el SOAP
                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                //Se define el ambiente del Dropdown de Environment
                var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);


                var i = TitleRow01 + 1;
                while ("" + _cells.GetCell(1, i).Value != "")
                    try
                    {
                        var wo = new WorkOrder();
                        wo.SetWorkOrderDto(_cells.GetEmptyIfNull(_cells.GetCell(1, i).Value2));
                        wo.requiredByDate = _cells.GetEmptyIfNull(_cells.GetCell(2, i).Value2);
                        wo.districtCode = DistrictCode;
                        try //actualizo la fecha
                        {
                            WorkOrderActions.ModifyWorkOrder(urlService, opSheet, wo);
                            _cells.GetCell(ResultColumn01 - 1, i).Value2 = "ACTUALIZADA";
                            _cells.GetCell(ResultColumn01 - 1, i).Style = StyleConstants.Success;
                        }
                        catch (Exception ex)
                        {
                            Debugger.LogError("RibbonEllipse:UpdateData()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                            _cells.GetCell(ResultColumn01 - 1, i).Value2 = ex.Message;
                            _cells.GetCell(ResultColumn01 - 1, i).Style = StyleConstants.Error;
                        }

                        try //finalizo la orden
                        {
                            WorkOrderActions.FinalizeWorkOrder(urlService, opSheet, wo);
                            _cells.GetCell(ResultColumn01, i).Value2 = "FINALIZADA";
                            _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                        }
                        catch (Exception ex)
                        {
                            Debugger.LogError("RibbonEllipse:UpdateData()",
                                "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                                ex.StackTrace);
                            _cells.GetCell(ResultColumn01, i).Value2 = ex.Message;
                            _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debugger.LogError("RibbonEllipse:UpdateData()",
                            "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" +
                            ex.StackTrace);
                        _cells.GetCell(ResultColumn01, i).Value2 = ex.Message;
                    }
                    finally
                    {
                        i++;
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debugger.LogError("RibbonEllipse:UpdateDate()",
                    "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
            }
            finally
            {
                _cells?.SetCursorDefault();
            }
        }

        private void GetWorkOrderStatusDataChangedValue(Range target, ListRanges changedRanges) //Excel.Range target)
        {
            try
            {
                if (target.Column != 1)
                    return;
                if (_cells.GetNullIfTrimmedEmpty(target.Text) == null)
                {
                    _cells.GetCell(target.Column, target.Row + 1).Value = "";
                    return;
                }

                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                var codeList = new List<string>(GetCompleteCodeList().Keys);
                var sqlQuery = Queries.GetWorkOrderStatusQuery(_eFunctions.DbReference, _eFunctions.DbLink, DistrictCode, codeList, WorkGroup, target.Text);
                var workOrderData = _eFunctions.GetQueryResult(sqlQuery);
                if (workOrderData == null) return;
                if (!workOrderData.IsClosed && workOrderData.HasRows)
                {
                    workOrderData.Read();
                    _cells.GetCell(3, target.Row).Value = workOrderData["CLIENTE"].ToString().Trim();
                    _cells.GetCell(4, target.Row).Value = workOrderData["CODESTADO"].ToString().Trim();
                    _cells.GetCell(5, target.Row).Value = workOrderData["ESTADOOT"].ToString().Trim();
                    _cells.GetCell(6, target.Row).Value = workOrderData["FIRMACLIENTE"].ToString().Trim();
                    _cells.GetCell(7, target.Row).Value = workOrderData["CALIFICACION"].ToString().Trim();
                    _cells.GetCell(8, target.Row).Value = workOrderData["NOCONFORME"].ToString().Trim();
                    _cells.GetCell(9, target.Row).Value = workOrderData["VALESABIERTOS"].ToString().Trim();
                    _cells.GetCell(10, target.Row).Value = workOrderData["DEMORASABIERTAS"].ToString().Trim();
                    _cells.GetCell(11, target.Row).Value = workOrderData["REQ_BY_DATE"].ToString().Trim();
                    _cells.GetCell(12, target.Row).Value =
                        workOrderData["FINAL_COSTS"].ToString().Trim().Equals("Y")
                            ? "FINALIZADA"
                            : "";
                }
                else
                {
                    _cells.GetCell(3, target.Row).Value = "";
                    _cells.GetCell(4, target.Row).Value = "";
                    _cells.GetCell(5, target.Row).Value = "";
                    _cells.GetCell(6, target.Row).Value = "";
                    _cells.GetCell(7, target.Row).Value = "";
                    _cells.GetCell(8, target.Row).Value = "";
                    _cells.GetCell(9, target.Row).Value = "";
                    _cells.GetCell(10, target.Row).Value = "";
                    _cells.GetCell(11, target.Row).Value = "";
                    _cells.GetCell(12, target.Row).Value = "";
                    _cells.GetCell(2, target.Row).Value = "ORDEN NO ENCONTRADA";
                }
            }
            catch (NullReferenceException)
            {
                _cells.GetCell(2, target.Row).Value = "NO FUE POSIBLE OBTENER INFORMACIÓN";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
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

    public static class Queries
    {
        public static string GetWorkOrderStatusQuery(string dbReference, string dbLink, string districtCode,
            List<string> jobDurCodesList, string workGroup, string workOrder)
        {
            var jobCodes = jobDurCodesList.Aggregate("", (current, jc) => current + "'" + jc + "', ");

            jobCodes = jobCodes.Substring(0, jobCodes.Length - 2);

            var sqlQuery = "" +
                           " SELECT" +
                           "   WO.WORK_ORDER WORK_ORDER," +
                           "   WO.WO_DESC CLIENTE," +
                           "   TRIM( DESCR_ORDEN.STD_VOLAT_2 ) || ' ' || TRIM( DESCR_ORDEN.STD_VOLAT_3 ) || ' ' || TRIM( DESCR_ORDEN.STD_VOLAT_4 ) || ' ' || TRIM( DESCR_ORDEN.STD_VOLAT_5 ) DESCRIPCION," +
                           "   WO.WO_STATUS_M CODESTADO," +
                           "   DECODE( TRIM( WO.WO_STATUS_M ), 'C', 'OK', 'NO' ) ESTADOOT," +
                           "   DECODE( TRIM( WO.PLAN_STR_DATE ), NULL, 'NO', 'OK' ) FIRMACLIENTE," +
                           "   DECODE( TRIM( REFERENCE_CODE.CALF ), NULL, 'NO', 'OK' ) CALIFICACION," +
                           "   DECODE( TRIM( REFERENCE_CODE.CONFORMIDAD ), 'Y', 'NO', NULL, 'OK' ) NOCONFORME," +
                           "   DECODE(" +
                           "     (SELECT COUNT( * ) FROM " + dbReference + ".MSF620" + dbLink + " WOV" +
                           "     INNER JOIN " + dbReference + ".MSF232" + dbLink +
                           "  RA ON WOV.WORK_ORDER                    = RA.WORK_ORDER" +
                           "     AND WOV.DSTRCT_CODE                                                = RA.DSTRCT_CODE" +
                           "     INNER JOIN " + dbReference + ".MSF140" + dbLink +
                           "  VM ON SUBSTR( RA.REQUISITION_NO, 1, 6 ) = VM.IREQ_NO" +
                           "     AND WOV.DSTRCT_CODE                                                = VM.DSTRCT_CODE WHERE WOV.DSTRCT_CODE = '" +
                           districtCode + "'" +
                           "     AND WOV.WORK_GROUP                                                 = '" + workGroup +
                           " '" +
                           "     AND WOV.WORK_ORDER                                                 = WO.WORK_ORDER" +
                           "     AND VM.AUTHSD_STATUS                                               = 'A'" +
                           "     AND VM.HDR_140_STATUS                                             IN( '0', '1' )" +
                           "     ) , '0', 'OK', 'NO' ) VALESABIERTOS," +
                           "   DECODE((" +
                           "     SELECT COUNT( * ) FROM (" +
                           "       SELECT DISTINCT" +
                           "         WO.WO_DESC CONTACTO," +
                           "         TRIM( REPLACE( TRIM( COM.STD_VOLAT_1 ) || ' ' || TRIM( COM.STD_VOLAT_2 ) || ' ' || TRIM( COM.STD_VOLAT_3 ) || ' ' || TRIM( COM.STD_VOLAT_4 ) || ' ' || TRIM( COM.STD_VOLAT_5 ), '.HEADING', '' ) ) DESCRIPCION," +
                           "         DUR.JOB_DUR_DATE," +
                           "         DUR.JOB_DUR_CODE," +
                           "         DUR.JOB_DUR_HOURS," +
                           "         LEAD( DUR.JOB_DUR_DATE ) OVER( PARTITION BY DUR.DSTRCT_CODE, WO.WORK_ORDER, DUR.JOB_DUR_CODE ORDER BY DUR.JOB_DUR_DATE ) LEAD_DATE," +
                           "         LEAD( DUR.JOB_DUR_HOURS ) OVER( PARTITION BY DUR.DSTRCT_CODE, WO.WORK_ORDER, DUR.JOB_DUR_CODE ORDER BY DUR.JOB_DUR_DATE ) LEAD_HOURS," +
                           "         CODE.TABLE_DESC," +
                           "         WO.WORK_ORDER," +
                           "         DUR.WORK_ORDER AS WORK_ORDER1," +
                           "         WO.RAISED_DATE AS APERTURA," +
                           "         WO.AUTHSD_DATE," +
                           "         WO.WO_TYPE," +
                           "         WO.WO_TYPE AS WO_TYPE1," +
                           "         WO.EQUIP_NO," +
                           "         WO.FINAL_COSTS," +
                           "         WO.WO_STATUS_M," +
                           "         WO.COMPLETED_CODE," +
                           "         WO.WORK_GROUP," +
                           "         WO.REQ_BY_DATE" +
                           "       FROM " + dbReference + ".MSF620" + dbLink + " WO" +
                           "       LEFT OUTER JOIN " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " COM" +
                           "       ON COM.STD_KEY = WO.DSTRCT_WO" +
                           "       LEFT JOIN " + dbReference + ".MSF622" + dbLink + " DUR" +
                           "       ON WO.DSTRCT_CODE  = DUR.DSTRCT_CODE" +
                           "       AND DUR.WORK_ORDER = WO.WORK_ORDER" +
                           "       LEFT JOIN " + dbReference + ".MSF010" + dbLink + " CODE" +
                           "       ON DUR.JOB_DUR_CODE     = CODE.TABLE_CODE" +
                           "       WHERE DUR.JOB_DUR_CODE IN( " + jobCodes + ")" +
                           "       AND WO.WORK_GROUP       = '" + workGroup + "'" +
                           "       AND WO.DSTRCT_CODE      = '" + districtCode + "'" +
                           "       AND COM.STD_TEXT_CODE   = 'WO'" +
                           "       AND CODE.TABLE_TYPE     = 'JI'" +
                           "       AND COM.STD_LINE_NO     = '0000'" +
                           "       ) X WHERE TRIM( X.JOB_DUR_CODE )" +
                           "       || ' - '" +
                           "       || TRIM( X.TABLE_DESC ) LIKE '%C%'" +
                           "     AND TRIM( X.JOB_DUR_DATE ) IS NOT NULL" +
                           "     AND TRIM( X.LEAD_DATE )    IS NULL" +
                           "     AND X.COMPLETED_CODE       <> '08'" +
                           "     AND X.JOB_DUR_HOURS         = 0" +
                           "     AND X.WORK_ORDER            = WO.WORK_ORDER" +
                           "     ) , '0', 'OK', 'NO' ) DEMORASABIERTAS," +
                           "     WO.REQ_BY_DATE," +
                           "     WO.FINAL_COSTS" +
                           " FROM " + dbReference + ".MSF620" + dbLink + " WO" +
                           " LEFT JOIN" +
                           "   (SELECT REF_TABLE.WORK_ORDER," +
                           "     MAX( DECODE( REF_TABLE.REF_NO, '016', REF_TABLE.REF_CODE, NULL ) ) CONFORMIDAD," +
                           "     MAX( DECODE( REF_TABLE.REF_NO, '021', REF_TABLE.REF_CODE, NULL ) ) CALF" +
                           "   FROM" +
                           "     (SELECT RC.REF_NO," +
                           "       RC.SCREEN_LITERAL," +
                           "       SUBSTR( RCD.ENTITY_VALUE, 6, 8 ) WORK_ORDER," +
                           "       RCD.REF_CODE REF_CODE," +
                           "       RCD.LAST_MOD_DATE || ' ' || RCD.LAST_MOD_TIME || ' ' || RCD.SEQ_NUM FECHA, TRIM( COM.STD_VOLAT_1 ) || ' ' || TRIM( COM.STD_VOLAT_2 ) || ' ' || TRIM( COM.STD_VOLAT_3 ) || ' ' || TRIM( COM.STD_VOLAT_4 ) || ' ' || TRIM( COM.STD_VOLAT_5 ) COM," +
                           "       MAX( RCD.LAST_MOD_DATE || ' ' || RCD.LAST_MOD_TIME || ' ' || RCD.SEQ_NUM ) OVER( PARTITION BY RCD.REF_NO, RCD.ENTITY_VALUE ) MAX_FECHA" +
                           "     FROM " + dbReference + ".MSF071" + dbLink + " RCD" +
                           "     INNER JOIN " + dbReference + ".MSF070" + dbLink + " RC" +
                           "     ON RCD.ENTITY_TYPE = RC.ENTITY_TYPE" +
                           "     AND RC.REF_NO      = RCD.REF_NO" +
                           "     LEFT JOIN " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " COM" +
                           "     ON COM.STD_KEY                       = RCD.STD_TXT_KEY" +
                           "     AND COM.STD_LINE_NO                  = '0000'" +
                           "     AND COM.STD_TEXT_CODE                = 'RC'" +
                           "     WHERE RCD.ENTITY_TYPE                = 'WKO'" +
                           "     AND RCD.REF_NO                      IN( '016', '021' )" +
                           "     AND SUBSTR( RCD.ENTITY_VALUE, 2, 4 ) = '" + districtCode + "'" +
                           "     ) REF_TABLE" +
                           "   WHERE REF_TABLE.FECHA = REF_TABLE.MAX_FECHA" +
                           "   GROUP BY REF_TABLE.WORK_ORDER" +
                           "   ) REFERENCE_CODE ON WO.WORK_ORDER = REFERENCE_CODE.WORK_ORDER" +
                           " LEFT JOIN " + dbReference + ".MSF096_STD_VOLAT" + dbLink + " DESCR_ORDEN" +
                           " ON WO.DSTRCT_WO               = DESCR_ORDEN.STD_KEY" +
                           " AND DESCR_ORDEN.STD_TEXT_CODE = 'WO'" +
                           " AND DESCR_ORDEN.STD_LINE_NO   = '0000'" +
                           " WHERE WO.WORK_ORDER           = '" + workOrder + "'" +
                           " AND WO.DSTRCT_CODE            = '" + districtCode + "'";

            return sqlQuery;
        }
    }
}