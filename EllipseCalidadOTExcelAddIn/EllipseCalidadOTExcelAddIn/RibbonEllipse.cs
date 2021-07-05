﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
//using WorkOrderService = EllipseWorkOrdersClassLibrary.WorkOrderService;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using System.Windows.Forms;
//using EllipseCommonsClassLibrary;
//using EllipseCommonsClassLibrary.Connections;
//using EllipseCommonsClassLibrary.Classes;
//using EllipseCommonsClassLibrary.Constants;
using EllipseWorkOrdersClassLibrary;
using WorkOrderService = EllipseWorkOrdersClassLibrary.WorkOrderService;
using Excel = Microsoft.Office.Interop.Excel;
//using Application = Microsoft.Office.Interop.Excel.Application;
//using FormAuthenticate = EllipseCommonsClassLibrary.FormAuthenticate;

using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Vsto.Excel;
using Debugger = SharedClassLibrary.Utilities.Debugger;
using SharedClassLibrary.Classes;
using System.Web.Services.Ellipse;
using System.Web.Services;
using SharedClassLibrary.Ellipse.Constants;

namespace EllipseCalidadOTExcelAddIn
{
    public partial class RibbonEllipse
    {

        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions;
        SharedClassLibrary.Ellipse.Forms.FormAuthenticate _frmAuth;
        //Application _excelApp;
        private Excel.Application _excelApp;

        private const string SheetName01 = "WorkOrders";
        private const int TitleRow01 = 7;
        private const int ResultColumn01 = 20;
        private const string TableName01 = "WorkOrderTable";
        private const string ValidationSheetName = "ValidationSheetWorkOrder";
        public string WoCode7 = "";
        //public bool CR;

        private Thread _thread;
        //private bool _progressUpdate = true;

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        private void bFormatear_Click(object sender, RibbonControlEventArgs e)
        {
            FormatSheet();

        }

        public void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new SharedClassLibrary.Ellipse.Forms.FormAuthenticate();
            _excelApp = Globals.ThisAddIn.Application;
            _excelApp.EnableEvents = true;
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

        private void FormatSheet()
        {
            try
            {
                //_excelApp = Globals.ThisAddIn.Application;
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

                //region CONSTRUYO LA HOJA 1
                _excelApp.Workbooks.Add();
                while (_excelApp.ActiveWorkbook.Sheets.Count < 3)
                    _excelApp.ActiveWorkbook.Worksheets.Add();
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                _cells.SetCursorWait();

                _excelApp.ActiveWorkbook.ActiveSheet.Name = SheetName01;
                _cells.CreateNewWorksheet(ValidationSheetName);//hoja de validación

                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "B2");

                _cells.GetCell("C1").Value = "CALIDAD DE INFORMACION DE WORK ORDERS - ELLIPSE 8";
                _cells.GetCell("C1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("C1", "J2");

                _cells.GetCell("L1").Value = "OBLIGATORIO";
                _cells.GetCell("L1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell("L2").Value = "OPCIONAL";
                _cells.GetCell("L2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("L3").Value = "INFORMATIVO";
                _cells.GetCell("L3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell("L4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell("L4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell("L5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell("L5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                var workGroupList = Groups.GetWorkGroupList().Select(g => g.Name).ToList();

                _cells.GetCell("A3").Value = "DISTRITO";
                _cells.GetCell("B3").Value = "ICOR";
                _cells.GetCell("A4").Value = "PROCESO";
                _cells.SetValidationList(_cells.GetCell("B4"), new List<string> { "ACARREO ELECTRICO", "ACARREO MECANICO", "CARGUE", "CARGUE ELECTRICO", "EALL", "PERFORACION Y VOLADURA", "TOR" });
                _cells.GetRange("A3", "A4").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetRange("B3", "B4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("G3").Value = "MAXIMO ORDENES";
                _cells.GetCell("G3").AddComment("AGREGAR MAXIMO DE ORDENES A MOSTRAR");
                _cells.GetCell("G3").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.SetValidationList(_cells.GetCell("E4"), new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50" });
                _cells.GetRange("G4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("H3").Value = "WORK GROUP";
                _cells.GetCell("H3").AddComment("BUSCAR POR GRUPO DE TRABAJO");
                _cells.GetCell("H3").Style = _cells.GetStyle(StyleConstants.TitleOptional);
               // _cells.SetValidationList(_cells.GetCell("F4"), workGroupList, ValidationSheetName, 3, false);
                _cells.SetValidationList(_cells.GetCell("H4"), new List<string> { "TANQ777", "TRACLLA", "VIAS", "ORUGAS", "CARGUE2", "PHIDCAS", "PHS", "EH320", "CAT2401", "CAT789C","K930E-4" });
                _cells.GetRange("H4").Style = _cells.GetStyle(StyleConstants.Select);

                _cells.GetCell("I3").Value = "EQUIPO";
                _cells.GetCell("I3").AddComment("BUSCAR POR EQUIPO");
                _cells.GetCell("I3").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetRange("I4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetCell("I4").NumberFormat = "@";

                _cells.GetCell("J3").Value = "WO_STATUS_M";
                _cells.GetCell("J3").AddComment("BUSCAR POR WO_STATUS_M");
                _cells.GetCell("J3").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell("K3").Value = "ASSIGN_TO";
                _cells.GetCell("K3").AddComment("BUSCAR POR ASSIGN_PERSON");
                _cells.GetCell("K3").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                var statusList = WoStatusList.GetStatusNames(true);
                _cells.SetValidationList(_cells.GetCell("J4"), statusList, ValidationSheetName, 4);
                //_cells.SetValidationList(_cells.GetCell("H4"), new List<string> { "'C'", "'O';'A'", "'C';'O';'A'" });
                _cells.GetCell("J4").NumberFormat = "@";
                _cells.GetCell("K4").NumberFormat = "@";
                _cells.GetRange("A1","A10000").NumberFormat = "@";
                _cells.GetRange("J4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetRange("K4").Style = _cells.GetStyle(StyleConstants.Select);


                _cells.GetCell("C3").Value = "FECHA DESDE";
                _cells.GetCell("D3").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + "01";
                _cells.GetCell("D3").AddComment("YYYYMMDD");
                _cells.GetCell("C4").Value = "FECHA HASTA";
                _cells.GetCell("D4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                _cells.GetCell("D4").AddComment("YYYYMMDD");
                _cells.GetRange("E3", "E4").Merge();
                _cells.GetCell("E3").Value = "Tipo Fecha";
                _cells.SetValidationList(_cells.GetCell("F3"), new List<string> { "PLAN_STR_DATE", "CLOSED_DT", "RAISED_DATE", "CREATION_DATE" });
                //_cells.GetCell("F3").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + "01";
                _cells.GetRange("F3", "F4").Merge();
                _cells.GetCell("F3").AddComment("SELECCIONA EL TIPO DE FECHA QUE SE VA A UTILIZAR PARA EJECUTAR LA CONSULTA");
                _cells.GetCell("F3").Value = "CLOSED_DT";
                //_cells.GetCell("E4").Value = "FECHA HASTA(CLOSED_DT)";
                //_cells.GetCell("F4").Value = string.Format("{0:0000}", DateTime.Now.Year) + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day);
                // _cells.GetCell("F4").AddComment("YYYYMMDD");
                _cells.GetRange("C3", "C4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("D3", "D4").Style = _cells.GetStyle(StyleConstants.Select);
                _cells.GetRange("E3", "E4").Style = _cells.GetStyle(StyleConstants.Option);
                _cells.GetRange("F3", "F4").Style = _cells.GetStyle(StyleConstants.Select);



                //_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01).Style = StyleConstants.TitleOptional;                

                //GENERAL
                _cells.GetCell(1, TitleRow01).Value = "WORK_ORDER";

                _cells.GetCell(2, TitleRow01).Value = "WO_STATUS_M";
                _cells.GetCell(3, TitleRow01).Value = "WO_DESC";
                _cells.GetCell(4, TitleRow01).Value = "EQUIP_NO";
                _cells.GetCell(5, TitleRow01).Value = "FLOTA";
                _cells.GetCell(6, TitleRow01).Value = "WORK_GROUP";
                _cells.GetCell(7, TitleRow01).Value = "LABOR ESTIMADA";
                
                _cells.GetCell(8, TitleRow01).Value = "DURACION";
                _cells.GetCell(9, TitleRow01).Value = "ASSIGN_PERSON";
                _cells.GetCell(10, TitleRow01).Value = "HORAS LABOR";
                _cells.GetCell(11, TitleRow01).Value = "COSTO MATERIAL REAL";
                _cells.GetCell(12, TitleRow01).Value = "FALLA FUNCIONAL";
                _cells.GetCell(13, TitleRow01).Value = "PARTE QUE FALLO";
                _cells.GetCell(14, TitleRow01).Value = "MODO DE FALLA";
                _cells.GetCell(15, TitleRow01).Value = "WO_JOB_CODEX8";
                _cells.GetCell(16, TitleRow01).Value = "COMENTARIO DE CIERRE";
                _cells.GetCell(17, TitleRow01).Value = "CALIFICACION";
                _cells.GetCell(17, TitleRow01).AddComment("<60% - Calidad Baja\n" +
                    ">=60% y <80% - Calidad Regular\n" +
                    ">=80% y <=99% - Calidad Buena\n" +
                    "100% - Calidad Excelente");
                _cells.GetCell(18, TitleRow01).Value = "GARANTIA";
                _cells.GetCell(19, TitleRow01).Value = "CALIFICADO_POR";
                _cells.GetCell(16, TitleRow01).Style = StyleConstants.TitleRequired;
                _cells.SetValidationList(_cells.GetCell(17, TitleRow01 + 1), new List<string> { "1 - BAJA", "2 - REGULAR", "3 - BUENA", "4 - EXCELENTE" });
                _cells.SetValidationList(_cells.GetCell(18, TitleRow01 + 1), new List<string> { "Y", "N"});

                _cells.GetCell(ResultColumn01, TitleRow01).Value = "RESULTADO";
                _cells.GetCell(ResultColumn01, TitleRow01).Style = StyleConstants.TitleResult;
                //_cells.GetRange(1, TitleRow01 + 1, ResultColumn01, TitleRow01 + 1).NumberFormat = NumberFormatConstants.Text;
                _cells.FormatAsTable(_cells.GetRange(1, TitleRow01, ResultColumn01, TitleRow01 + 1), TableName01);
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                               

                _excelApp.ActiveWorkbook.Sheets[1].Select(Type.Missing);
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:FormatSheet()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error al intentar crear el encabezado de la hoja. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
            }
        }

        private void Consultar()
        {
            try
            {
                
                var currentRow = TitleRow01 + 1;
                var sqlQuery = @"WITH A AS
                                (
                                  SELECT
                                    W.WORK_ORDER,
                                    W.WO_STATUS_M,
                                    W.WO_DESC,
                                    W.WO_TYPE,
                                    W.MAINT_TYPE,
                                    SUBSTR(W.DSTRCT_ACCT_CODE, 5, LENGTH(W.DSTRCT_ACCT_CODE)) AS CENTRO,
                                    W.EQUIP_NO,
                                    EQ.FLOTA_ELLIPSE,
                                    EQ.PROCESO,
                                    W.WORK_GROUP,
                                    W.REQ_START_DATE,
                                    W.PLAN_STR_DATE,
                                    W.WO_JOB_CODEX1,
                                    W.WO_JOB_CODEX2,
                                    W.WO_JOB_CODEX3,
                                     (CASE WHEN W.PLAN_STR_DATE <> ' ' THEN 
                                      (CASE 
                                      WHEN TO_CHAR(TO_DATE(W.PLAN_STR_DATE, 'YYYYMMDD'), 'D') <= 3 THEN TO_CHAR(TO_DATE(W.PLAN_STR_DATE, 'YYYYMMDD') - 6, 'YYYYWW')
                                      ELSE TO_CHAR(TO_DATE(W.PLAN_STR_DATE, 'YYYYMMDD'), 'YYYYWW')
                                      END)  
                                    ELSE '' END )SEMANA_PLAN,
                                    W.ORIG_PRIORITY,  
                                    E.ACT_DUR_HRS AS DURACION,  
                                    E.ACT_MAT_COST AS COSTOS_MAT,  
                                    E.ACT_LAB_HRS AS HORAS_LAB,  
                                    E.EST_LAB_HRS AS LABOR,  
                                    E.CALC_LAB_HRS AS LABOR_CAL,  
                                    W.RELATED_WO,  
                                    W.ASSIGN_PERSON, 
                                    SIGMAN.FNU_COMENTARIOS_OT2@DBLSIG(W.WORK_ORDER,'CW','1') AS COMENTARIO_CIERRE,
                                    W.WO_JOB_CODEX8
                                FROM
                                    ELLIPSE.MSF620 W 
                                    INNER JOIN ELLIPSE.MSF621 E ON(W.WORK_ORDER = E.WORK_ORDER)  
                                    INNER JOIN SIGMAN.EQMTLIST@DBLSIG EQ ON(RPAD(EQ.EQU,12,' ') = W.EQUIP_NO)
                                WHERE
                                    W.DSTRCT_CODE = 'ICOR' " + @"
                                    AND W."  + _cells.GetCell("F3").Value +  @" BETWEEN '" + _cells.GetEmptyIfNull(_cells.GetCell("D3").Value) + "' AND '" + _cells.GetEmptyIfNull(_cells.GetCell("D4").Value) + "' ";
                
                //"AND W.WO_JOB_CODEX10 <> 'IG'" +
                sqlQuery += "AND SUBSTR(W.WORK_ORDER, 1, 2) <> 'EV'" +
                             "AND W.WORK_ORDER = E.WORK_ORDER " +
                            "AND EQ.EQU = TRIM(W.EQUIP_NO) ";

                if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("H4").Value) != null && _cells.GetNullIfTrimmedEmpty(_cells.GetCell("B4").Value) != null)
                {
                    //Datos = "";
                    //_cells.GetCell("A5").Value = "";
                    MessageBox.Show("Solo puede selecionar La busqueda por WORK GROUP o PROCESO");
                }
                else {


                    if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("H4").Value) != null)
                    {
                        sqlQuery += "AND W.WORK_GROUP = '" + _cells.GetEmptyIfNull(_cells.GetCell("H4").Value) + "' ";
                    }
                    if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("B4").Value) != null)
                    {
                        sqlQuery += "AND TRIM(EQ.PROCESO) = '" + _cells.GetEmptyIfNull(_cells.GetCell("B4").Value) + "' ";
                    }
                    if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("I4").Value) != null)
                    {
                        sqlQuery += "AND TRIM(W.EQUIP_NO) = '" + _cells.GetEmptyIfNull(_cells.GetCell("I4").Value) + "' ";
                    }
                    if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("J4").Value) != null)
                    {
                        //MessageBox.Show(WoStatusList.GetStatusCode(_cells.GetEmptyIfNull(_cells.GetCell("H4").Value)));
                        //sqlQuery += "AND W.WO_STATUS_M in ('" + _cells.GetEmptyIfNull( _cells.GetCell("H4").Value.Replace(';', ',')) + ") ";
                        sqlQuery += "AND W.WO_STATUS_M = '" + WoStatusList.GetStatusCode(_cells.GetEmptyIfNull(_cells.GetCell("J4").Value)) + "' ";
                    }
                    if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("K4").Value) != null)
                    {
                        //MessageBox.Show(WoStatusList.GetStatusCode(_cells.GetEmptyIfNull(_cells.GetCell("H4").Value)));
                        //sqlQuery += "AND W.WO_STATUS_M in ('" + _cells.GetEmptyIfNull( _cells.GetCell("H4").Value.Replace(';', ',')) + ") ";
                        sqlQuery += "AND W.ASSIGN_PERSON = RPAD('" + _cells.GetEmptyIfNull(_cells.GetCell("K4").Value) + "',10,' ') ";
                    }
                    if (WoCode7 != "")
                    {
                        sqlQuery += "AND TRIM(W.WO_JOB_CODEX8) = 'SC' ";
                        WoCode7 = "";
                    }

                    sqlQuery += @" ),
                            B AS
                            (
                              SELECT
                                A.WORK_ORDER,
                                A.WO_DESC,
                                A.WO_STATUS_M,
                                A.WO_TYPE,
                                A.MAINT_TYPE,
                                A.CENTRO,
                                A.EQUIP_NO,
                                A.FLOTA_ELLIPSE,
                                A.PROCESO,
                                A.WORK_GROUP,
                                A.REQ_START_DATE,
                                A.PLAN_STR_DATE,
                                A.SEMANA_PLAN,
                                A.ORIG_PRIORITY,
                                A.LABOR,
                                A.LABOR_CAL,
                                A.COSTOS_MAT,
                                A.HORAS_LAB,
                                A.DURACION,
                                A.RELATED_WO,
                                A.ASSIGN_PERSON,
                                LENGTH(TRIM(A.COMENTARIO_CIERRE)) AS T_COMENTARIO,
                                A.COMENTARIO_CIERRE,
                                A.WO_JOB_CODEX1 AS FALLA_FUNCIONAL,
                                A.WO_JOB_CODEX2 AS PARTE_FALLO,
                                A.WO_JOB_CODEX3 AS MODO_FALLA,
                                A.WO_JOB_CODEX8,
                                (
                                  SELECT
                                  CASE WHEN TO_NUMBER(REF_CODE) = '1' THEN '1 - BAJA'
                                       WHEN TO_NUMBER(REF_CODE) = '2' THEN '2 - REGULAR'
                                       WHEN TO_NUMBER(REF_CODE) = '3' THEN '3 - BUENA'
                                       WHEN TO_NUMBER(REF_CODE) = '4' THEN '4 - EXCELENTE'
                                       ELSE ''
                                  END CALIDAD
                                    FROM
                                      ELLIPSE.MSF071 RC,
                                      ELLIPSE.MSF070 RCE
                                    WHERE
                                      RC.ENTITY_TYPE = RCE.ENTITY_TYPE
                                      AND RC.REF_NO = RCE.REF_NO
                                      AND RCE.ENTITY_TYPE = 'WKO'
                                      AND RC.REF_NO = '034'
                                      AND RC.SEQ_NUM = '001'
                                      AND SUBSTR(RC.ENTITY_VALUE, 6, 8) = A.WORK_ORDER
                                ) AS CALIDAD,
                                (
                                  SELECT
                                    REF_CODE
                                    FROM
                                      ELLIPSE.MSF071 RC,
                                      ELLIPSE.MSF070 RCE
                                    WHERE
                                      RC.ENTITY_TYPE = RCE.ENTITY_TYPE
                                      AND RC.REF_NO = RCE.REF_NO
                                      AND RCE.ENTITY_TYPE = 'WKO'
                                      AND RC.REF_NO = '024'
                                      AND RC.SEQ_NUM = '001'
                                      AND SUBSTR(RC.ENTITY_VALUE, 6, 8) = A.WORK_ORDER
                                ) AS GARANTIA,
                                (
                                  SELECT
                                    TRIM(REF_CODE)
                                    FROM
                                      ELLIPSE.MSF071 RC,
                                      ELLIPSE.MSF070 RCE
                                    WHERE
                                      RC.ENTITY_TYPE = RCE.ENTITY_TYPE
                                      AND RC.REF_NO = RCE.REF_NO
                                      AND RCE.ENTITY_TYPE = 'WKO'
                                      AND RC.REF_NO = '035'
                                      AND RC.SEQ_NUM = '001'
                                      AND SUBSTR(RC.ENTITY_VALUE, 6, 8) = A.WORK_ORDER
                                ) AS CALIFICADO_POR
                              FROM
                                A
                            )
                            SELECT
                              B.WORK_ORDER,  
                              B.WO_DESC, 
                              B.WO_STATUS_M, 
                              B.WO_TYPE,  
                              B.MAINT_TYPE,  
                              B.CENTRO,  
                              B.EQUIP_NO,  
                              B.FLOTA_ELLIPSE,  
                              B.PROCESO,  
                              B.WORK_GROUP,  
                              B.REQ_START_DATE,  
                              B.PLAN_STR_DATE,  
                              B.SEMANA_PLAN,  
                              B.ORIG_PRIORITY,  
                              B.LABOR,  
                              B.LABOR_CAL,  
                              B.DURACION,  
                              B.ASSIGN_PERSON,  
                              B.FALLA_FUNCIONAL,  
                              B.PARTE_FALLO,  
                              B.MODO_FALLA,  
                              B.RELATED_WO,  
                              B.COSTOS_MAT,  
                              B.HORAS_LAB,  
                              B.T_COMENTARIO,  
                              B.WO_JOB_CODEX8,
                              B.COMENTARIO_CIERRE,  
                              B.CALIDAD,
                              B.GARANTIA,
                              B.CALIFICADO_POR
                            FROM
                              B WHERE B.CALIDAD IS NULL ";
                    /*if (WoCode7 != "")
                    {
                        WoCode7 = "";
                        if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("G4").Value) != null)
                        {
                            sqlQuery += " AND rownum <= '" + _cells.GetEmptyIfNull(_cells.GetCell("G4").Value) + "' ";
                        }
                    }
                    else */
                    if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("G4").Value) != null)
                    {
                        sqlQuery += " AND rownum <= '" + _cells.GetEmptyIfNull(_cells.GetCell("G4").Value) + "' ";
                    }

                    _cells.GetCell("A5").Value = "Consultando Informacion. Por favor espere...";

                    //var dbi = Environments.GetDatabaseItem(/*drpEnvironment.SelectedItem.Label*/Environments.EllipseSigmanProductivo);
                    //dbi.DbPassword = "ventyx15";
                    //dbi.DbUser = "consulbo";

                    //dbi.DbEncodedPassword = "TPZPyIEoE7gfD6TZUaKys4yxQWTAe5BNWK1wjmjo1CVdYnbiAzyhOPSszjHNvUWPSWH4cq9q2Cs4gDdGCK7+JfgpqjmYXTF+8VSfW78zIcPJafyHOtwBweS+QjZEFa9W";
                    //_eFunctions.SetDBSettings(dbi.DbName, dbi.DbUser, dbi.DbPassword);
                    _eFunctions.SetDBSettings(Environments.EllipseSigmanProductivo);
                    _eFunctions.SetConnectionTimeOut(0);
                    var odr = _eFunctions.GetQueryResult(sqlQuery);
                    _cells.ClearTableRange(TableName01);
                    _cells.GetRange(TableName01).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    while (odr.Read())
                    {
                        _cells.GetCell("A" + currentRow).NumberFormat = "@";
                        _cells.GetCell("A" + currentRow).Value = odr["WORK_ORDER"] + "";
                        _cells.GetCell("B" + currentRow).Value = odr["WO_STATUS_M"] + "";
                        _cells.GetCell("C" + currentRow).Value = odr["WO_DESC"] + "";
                        _cells.GetCell("D" + currentRow).Value = odr["EQUIP_NO"] + "";
                        _cells.GetCell("E" + currentRow).Value = odr["FLOTA_ELLIPSE"] + "";
                        _cells.GetCell("F" + currentRow).Value = odr["WORK_GROUP"] + "";
                        _cells.GetCell("G" + currentRow).Value = odr["LABOR"];
                        _cells.GetCell("H" + currentRow).Value = odr["DURACION"] + "";
                        _cells.GetCell("I" + currentRow).Value = odr["ASSIGN_PERSON"] + "";
                        _cells.GetCell("J" + currentRow).Value = odr["HORAS_LAB"] + "";
                        _cells.GetCell("K" + currentRow).Value = odr["COSTOS_MAT"] + "";
                        _cells.GetCell("L" + currentRow).Value = odr["FALLA_FUNCIONAL"] + "";
                        _cells.GetCell("M" + currentRow).Value = odr["PARTE_FALLO"] + "";
                        _cells.GetCell("N" + currentRow).Value = odr["MODO_FALLA"] + "";
                        _cells.GetCell("O" + currentRow).Value = odr["WO_JOB_CODEX8"] + "";
                        _cells.GetCell("P" + currentRow).Value = odr["COMENTARIO_CIERRE"] + "";
                        _cells.GetCell("P" + currentRow).EntireColumn.ColumnWidth = 150;
                        _cells.GetCell("P" + currentRow).WrapText = true;
                        _cells.GetCell("Q" + currentRow).Value = odr["CALIDAD"] + "";
                        _cells.GetCell("Q" + currentRow).NumberFormat = "###,##%";
                        _cells.GetCell("R" + currentRow).Value = odr["GARANTIA"] + "";
                        _cells.GetCell("S" + currentRow).Value = odr["CALIFICADO_POR"] + "";

                        if (Convert.ToDouble(odr["HORAS_LAB"]) <= 0)
                        {
                            _cells.GetCell("J" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            _cells.GetCell("J" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        if (Convert.ToDouble(odr["DURACION"]) <= 0)
                        {
                            _cells.GetCell("H" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                        }
                        else
                        {
                            _cells.GetCell("H" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }


                        if (Convert.ToDouble(odr["COSTOS_MAT"]) <= 0)
                        {

                            _cells.GetCell("K" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            _cells.GetCell("K" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }



                        if (odr["FALLA_FUNCIONAL"] == null || odr["FALLA_FUNCIONAL"].ToString().Trim() == "")
                        {
                            _cells.GetCell("L" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            _cells.GetCell("L" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        if (odr["PARTE_FALLO"] == null || odr["PARTE_FALLO"].ToString().Trim() == "")
                        {
                            _cells.GetCell("M" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            _cells.GetCell("M" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        if (odr["MODO_FALLA"] == null || odr["MODO_FALLA"].ToString().Trim() == "")
                        {
                            _cells.GetCell("N" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            _cells.GetCell("N" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        if (odr["COMENTARIO_CIERRE"] == null || odr["COMENTARIO_CIERRE"].ToString().Trim() == "")
                        {
                            _cells.GetCell("P" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            _cells.GetCell("P" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        currentRow = currentRow + 1;
                        //_cells.GetCell(1, currentRow).Select();
                        _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                    }


                }
                _cells.GetCell("A5").Value = "";
                //_cells.GetCell("E4").Value = "";
                //_cells.GetCell("F4").Value = "";

                MessageBox.Show("Consulta finalizada. No hay mas datos");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void Reconsultar()
        {
            try
            {

                var currentRow = TitleRow01 + 1;
                var sqlQuery = @" WITH A AS
                                (
                                  SELECT
                                    W.WORK_ORDER,
                                    W.WO_STATUS_M,
                                    W.WO_DESC,
                                    W.WO_TYPE,
                                    W.MAINT_TYPE,
                                    SUBSTR(W.DSTRCT_ACCT_CODE, 5, LENGTH(W.DSTRCT_ACCT_CODE)) AS CENTRO,
                                    W.EQUIP_NO,
                                    EQ.FLOTA_ELLIPSE,
                                    EQ.PROCESO,
                                    W.WORK_GROUP,
                                    W.REQ_START_DATE,
                                    W.PLAN_STR_DATE,
                                    W.WO_JOB_CODEX1,
                                    W.WO_JOB_CODEX2,
                                    W.WO_JOB_CODEX3,
                                     (CASE WHEN W.PLAN_STR_DATE <> ' ' THEN 
                                      (CASE 
                                      WHEN TO_CHAR(TO_DATE(W.PLAN_STR_DATE, 'YYYYMMDD'), 'D') <= 3 THEN TO_CHAR(TO_DATE(W.PLAN_STR_DATE, 'YYYYMMDD') - 6, 'YYYYWW')
                                      ELSE TO_CHAR(TO_DATE(W.PLAN_STR_DATE, 'YYYYMMDD'), 'YYYYWW')
                                      END)  
                                    ELSE '' END )SEMANA_PLAN,
                                    W.ORIG_PRIORITY,  
                                    E.ACT_DUR_HRS AS DURACION,  
                                    E.ACT_MAT_COST AS COSTOS_MAT,  
                                    E.ACT_LAB_HRS AS HORAS_LAB,  
                                    E.EST_LAB_HRS AS LABOR,  
                                    E.CALC_LAB_HRS AS LABOR_CAL,  
                                    W.RELATED_WO,  
                                    W.ASSIGN_PERSON, 
                                    SIGMAN.FNU_COMENTARIOS_OT2@DBLSIG(W.WORK_ORDER,'CW','1') AS COMENTARIO_CIERRE,
                                    W.WO_JOB_CODEX8
                                FROM
                                    ELLIPSE.MSF620 W 
                                    INNER JOIN ELLIPSE.MSF621 E ON(W.WORK_ORDER = E.WORK_ORDER)  
                                    INNER JOIN SIGMAN.EQMTLIST@DBLSIG EQ ON(RPAD(EQ.EQU,12,' ') = W.EQUIP_NO)
                                WHERE
                                    W.DSTRCT_CODE = 'ICOR'
                                    AND W.WO_JOB_CODEX10 <> 'IG'
                                    AND SUBSTR(W.WORK_ORDER, 1, 2) <> 'EV' ";

               
                if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("A8").Value) != null)
                {
                    sqlQuery += "AND W.WORK_ORDER in (";
                    sqlQuery += "'" + _cells.GetEmptyIfNull(_cells.GetCell("A8").Value + "'");
                    for (int i = 9; i < 1000; i++)
                    {
                        if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell("A" + i).Value) != null)
                        {
                            sqlQuery += ",'" + _cells.GetEmptyIfNull(_cells.GetCell("A" + i).Value) + "'";
                        }
                        else {
                            break;
                        }
                    }
                    sqlQuery += ") ";
                }

                sqlQuery += @" ),
                            B AS
                            (
                              SELECT
                                A.WORK_ORDER,
                                A.WO_DESC,
                                A.WO_STATUS_M,
                                A.WO_TYPE,
                                A.MAINT_TYPE,
                                A.CENTRO,
                                A.EQUIP_NO,
                                A.FLOTA_ELLIPSE,
                                A.PROCESO,
                                A.WORK_GROUP,
                                A.REQ_START_DATE,
                                A.PLAN_STR_DATE,
                                A.SEMANA_PLAN,
                                A.ORIG_PRIORITY,
                                A.LABOR,
                                A.LABOR_CAL,
                                A.COSTOS_MAT,
                                A.HORAS_LAB,
                                A.DURACION,
                                A.RELATED_WO,
                                A.ASSIGN_PERSON,
                                LENGTH(TRIM(A.COMENTARIO_CIERRE)) AS T_COMENTARIO,
                                A.COMENTARIO_CIERRE,
                                A.WO_JOB_CODEX1 AS FALLA_FUNCIONAL,
                                A.WO_JOB_CODEX2 AS PARTE_FALLO,
                                A.WO_JOB_CODEX3 AS MODO_FALLA,
                                A.WO_JOB_CODEX8,
(
                                  SELECT
                                  CASE WHEN TO_NUMBER(REF_CODE) = '1' THEN '1 - BAJA'
                                       WHEN TO_NUMBER(REF_CODE) = '2' THEN '2 - REGULAR'
                                       WHEN TO_NUMBER(REF_CODE) = '3' THEN '3 - BUENA'
                                       WHEN TO_NUMBER(REF_CODE) = '4' THEN '4 - EXCELENTE'
                                       ELSE ''
                                  END CALIDAD
                                    FROM
                                      ELLIPSE.MSF071 RC,
                                      ELLIPSE.MSF070 RCE
                                    WHERE
                                      RC.ENTITY_TYPE = RCE.ENTITY_TYPE
                                      AND RC.REF_NO = RCE.REF_NO
                                      AND RCE.ENTITY_TYPE = 'WKO'
                                      AND RC.REF_NO = '034'
                                      AND RC.SEQ_NUM = '001'
                                      AND SUBSTR(RC.ENTITY_VALUE, 6, 8) = A.WORK_ORDER
                                ) AS CALIDAD,
                                (
                                  SELECT
                                    REF_CODE
                                    FROM
                                      ELLIPSE.MSF071 RC,
                                      ELLIPSE.MSF070 RCE
                                    WHERE
                                      RC.ENTITY_TYPE = RCE.ENTITY_TYPE
                                      AND RC.REF_NO = RCE.REF_NO
                                      AND RCE.ENTITY_TYPE = 'WKO'
                                      AND RC.REF_NO = '024'
                                      AND RC.SEQ_NUM = '001'
                                      AND SUBSTR(RC.ENTITY_VALUE, 6, 8) = A.WORK_ORDER
                                ) AS GARANTIA,
                                (
                                  SELECT
                                    TRIM(REF_CODE)
                                    FROM
                                      ELLIPSE.MSF071 RC,
                                      ELLIPSE.MSF070 RCE
                                    WHERE
                                      RC.ENTITY_TYPE = RCE.ENTITY_TYPE
                                      AND RC.REF_NO = RCE.REF_NO
                                      AND RCE.ENTITY_TYPE = 'WKO'
                                      AND RC.REF_NO = '035'
                                      AND RC.SEQ_NUM = '001'
                                      AND SUBSTR(RC.ENTITY_VALUE, 6, 8) = A.WORK_ORDER
                                ) AS CALIFICADO_POR
                              FROM
                                A
                            )
                            SELECT
                              B.WORK_ORDER,  
                              B.WO_DESC, 
                              B.WO_STATUS_M, 
                              B.WO_TYPE,  
                              B.MAINT_TYPE,  
                              B.CENTRO,  
                              B.EQUIP_NO,  
                              B.FLOTA_ELLIPSE,  
                              B.PROCESO,  
                              B.WORK_GROUP,  
                              B.REQ_START_DATE,  
                              B.PLAN_STR_DATE,  
                              B.SEMANA_PLAN,  
                              B.ORIG_PRIORITY,  
                              B.LABOR,  
                              B.LABOR_CAL,  
                              B.DURACION,  
                              B.ASSIGN_PERSON,  
                              B.FALLA_FUNCIONAL,  
                              B.PARTE_FALLO,  
                              B.MODO_FALLA,  
                              B.RELATED_WO,  
                              B.COSTOS_MAT,  
                              B.HORAS_LAB,  
                              B.T_COMENTARIO, 
                              B.WO_JOB_CODEX8, 
                              B.COMENTARIO_CIERRE,  
                              B.CALIDAD,
                              B.GARANTIA,
                              B.CALIFICADO_POR
                            FROM
                              B /*WHERE B.CALIDAD IS NULL*/";
                

                _cells.GetCell("A5").Value = "Consultando Informacion. Por favor espere...";
                //var dbi = Environments.GetDatabaseItem(drpEnvironment.SelectedItem.Label);
                //dbi.DbUser = "consulbo";
                //dbi.DbEncodedPassword = "TPZPyIEoE7gfD6TZUaKys4yxQWTAe5BNWK1wjmjo1CVdYnbiAzyhOPSszjHNvUWPSWH4cq9q2Cs4gDdGCK7+JfgpqjmYXTF+8VSfW78zIcPJafyHOtwBweS+QjZEFa9W";
                //_eFunctions.SetDBSettings(dbi.DbName, dbi.DbUser, dbi.DbPassword);
                _eFunctions.SetDBSettings(Environments.EllipseSigmanProductivo);
                _eFunctions.SetConnectionTimeOut(0);
                var odr = _eFunctions.GetQueryResult(sqlQuery);
                _cells.ClearTableRange(TableName01);
                _cells.GetRange(TableName01).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                while (odr.Read())
                {
                    _cells.GetCell("A" + currentRow).Value = odr["WORK_ORDER"] + "";
                    _cells.GetCell("B" + currentRow).Value = odr["WO_STATUS_M"] + "";
                    _cells.GetCell("C" + currentRow).Value = odr["WO_DESC"] + "";
                    _cells.GetCell("D" + currentRow).Value = odr["EQUIP_NO"] + "";
                    _cells.GetCell("E" + currentRow).Value = odr["FLOTA_ELLIPSE"] + "";
                    _cells.GetCell("F" + currentRow).Value = odr["WORK_GROUP"] + "";
                    _cells.GetCell("G" + currentRow).Value = odr["LABOR"];
                    _cells.GetCell("H" + currentRow).Value = odr["DURACION"] + "";
                    _cells.GetCell("I" + currentRow).Value = odr["ASSIGN_PERSON"] + "";
                    _cells.GetCell("J" + currentRow).Value = odr["HORAS_LAB"] + "";
                    _cells.GetCell("K" + currentRow).Value = odr["COSTOS_MAT"] + "";
                    _cells.GetCell("L" + currentRow).Value = odr["FALLA_FUNCIONAL"] + "";
                    _cells.GetCell("M" + currentRow).Value = odr["PARTE_FALLO"] + "";
                    _cells.GetCell("N" + currentRow).Value = odr["MODO_FALLA"] + "";
                    _cells.GetCell("O" + currentRow).Value = odr["WO_JOB_CODEX8"] + "";
                    _cells.GetCell("P" + currentRow).Value = odr["COMENTARIO_CIERRE"] + "";
                    //_excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                    _excelApp.ActiveCell.WrapText = true;
                    _cells.GetCell("Q" + currentRow).Value = odr["CALIDAD"] + "";
                    _cells.GetCell("Q" + currentRow).NumberFormat = "###,##%";
                    _cells.GetCell("R" + currentRow).Value = odr["GARANTIA"] + "";
                    _cells.GetCell("S" + currentRow).Value = odr["CALIFICADO_POR"] + "";

                    if (Convert.ToDouble(odr["HORAS_LAB"]) <= 0)
                    {
                        _cells.GetCell("J" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        _cells.GetCell("J" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }

                    if (Convert.ToDouble(odr["DURACION"]) <= 0)
                    {
                        _cells.GetCell("H" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                    }
                    else
                    {
                        _cells.GetCell("H" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }


                    if (Convert.ToDouble(odr["COSTOS_MAT"]) <= 0)
                    {

                        _cells.GetCell("K" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        _cells.GetCell("K" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }



                    if (odr["FALLA_FUNCIONAL"] == null || odr["FALLA_FUNCIONAL"].ToString().Trim() == "")
                    {
                        _cells.GetCell("L" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        _cells.GetCell("L" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }

                    if (odr["PARTE_FALLO"] == null || odr["PARTE_FALLO"].ToString().Trim() == "")
                    {
                        _cells.GetCell("M" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        _cells.GetCell("M" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }

                    if (odr["MODO_FALLA"] == null || odr["MODO_FALLA"].ToString().Trim() == "")
                    {
                        _cells.GetCell("N" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        _cells.GetCell("N" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }

                    if (odr["COMENTARIO_CIERRE"] == null || odr["COMENTARIO_CIERRE"].ToString().Trim() == "")
                    {
                        _cells.GetCell("P" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                    else
                    {
                        _cells.GetCell("P" + currentRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }


                    currentRow = currentRow + 1;
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                _cells.GetCell("A5").Value = "";
                //_cells.GetCell("E4").Value = "";
                //_cells.GetCell("F4").Value = "";

                MessageBox.Show("Consulta finalizada. No hay mas datos");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse:GetQueryResult()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error. " + ex.Message);
            }
            finally
            {
                if (_cells != null) _cells.SetCursorDefault();
                _eFunctions.CloseConnection();
            }
        }

        private void bConsultar_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    
                    _thread = new Thread(Consultar);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

      



        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort(); _cells.GetCell("A5").Value = "";
                if (_cells != null) _cells.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
                _cells.GetCell("A5").Value = "";
            }
        }

        private void bLimpiar_Click(object sender, RibbonControlEventArgs e)
        {
            _cells.GetCell("A5").Value = "";
            _cells.GetCell("E4").Value = "";
            _cells.GetCell("F4").Value = "";
            _cells.ClearTableRange(TableName01);
            
            _cells.GetRange(TableName01).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
        }

        private void bCalificar_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                    _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                    if (_frmAuth.ShowDialog() != DialogResult.OK) return;
                    _thread = new Thread(CalificarOT);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }

        private void CalificarOT()
        {
            _eFunctions.SetDBSettings(Environments.EllipseProductivo);
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _cells.ClearTableRangeColumn(TableName01, ResultColumn01);

            var i = TitleRow01 + 1;
            var urlService = Environments.GetServiceUrl(Environments.EllipseProductivo);
            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);
            var district = _cells.GetNullIfTrimmedEmpty(_frmAuth.EllipseDstrct) ?? "ICOR";
            var userName = _frmAuth.EllipseUser.ToUpper();

            while (!string.IsNullOrEmpty(_cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2)))
            {
                try
                {
                    //if (_cells.GetNullIfTrimmedEmpty(_cells.GetCell((ResultColumn01-1), i).Value) == null)
                    //{
                        UpdateReferenceCodes(i, district, _cells.GetNullOrTrimmedValue(_cells.GetCell(1, i).Value2));
                        _cells.GetCell(ResultColumn01, i).Value = "OT CALIFICADA";
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Success;
                    //}
                    /*else
                    {
                        _cells.GetCell(ResultColumn01, i).Value = "OT NO CALIFICADA";
                        _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Warning;
                    }*/

                }
                catch (Exception ex)
                {
                    _cells.GetCell(ResultColumn01, i).Style = StyleConstants.Error;
                    _cells.GetCell(ResultColumn01, i).Value = "ERROR: " + ex.Message;
                    Debugger.LogError("RibbonEllipse.cs:CalificarOT()", ex.Message);
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

        public string Consulta()
        {
            string sqlQuery;

            return sqlQuery = @" '" + _cells.GetEmptyIfNull(_cells.GetCell("D3").Value) +  "' AND '" + _cells.GetEmptyIfNull(_cells.GetCell("D4").Value) +
                                        "' AND W.WO_STATUS_M = 'C'"+
                                        "AND W.WO_JOB_CODEX10 <> 'IG'"+
                                        "AND SUBSTR(W.WORK_ORDER, 1, 2) <> 'EV'"+
                                        "AND W.WORK_ORDER = E.WORK_ORDER "+
                                        "AND EQ.EQU = TRIM(W.EQUIP_NO)";
        }




        private void UpdateReferenceCodes(int fila, string distrito, string WO)
        {

            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _cells.SetCursorWait();

            _eFunctions.SetDBSettings(Environments.EllipseProductivo);
            var urlService = Environments.GetServiceUrl(Environments.EllipseProductivo);


            var opSheet = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost,
                maxInstances = 100,
                maxInstancesSpecified = true,
                returnWarnings = Debugger.DebugWarnings,
                returnWarningsSpecified = true
            };

            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            /*while (!string.IsNullOrEmpty("" + _cells.GetCell(1, fila).Value))
            {
                try
                {*/
            //GENERAL
            var district = distrito;
            var workOrder = WO;
            var calif = "";
            var calificacion = _cells.GetEmptyIfNull(_cells.GetCell(17, fila).Value);
            var Garantia = _cells.GetEmptyIfNull(_cells.GetCell(18, fila).Value);
            var CalifPor = _cells.GetEmptyIfNull(_cells.GetCell(19, fila).Value);
            //var CalificacionCalidadOT = "";
            //var CalificadoPor = "";

            if (calificacion == "1 - BAJA")
            {
                calif = "1";
            }
            else if (calificacion == "2 - REGULAR")
            {
                calif = "2";
            }
            else if (calificacion == "3 - BUENA")
            {
                calif = "3";
            }
            else if (calificacion == "4 - EXCELENTE")
            {
                calif = "4";
            }
            var User = "";
            if(calif != "")
            {
                //User = _frmAuth.EllipseUser;
                User = CalifPor;
            }

            var woRefCodes = new WorkOrderReferenceCodes
            {
                CalificacionCalidadOt = calif,
                CalificacionCalidadPor = User,
                Garantia = Garantia
            };
            var replyRefCode = WorkOrderActions.UpdateWorkOrderReferenceCodes(_eFunctions, urlService, opSheet, district, workOrder, woRefCodes);

            if (replyRefCode.Errors != null && replyRefCode.Errors.Length > 0)
            {
                var errorList = "";
                // ReSharper disable once LoopCanBeConvertedToQuery
                foreach (var error in replyRefCode.Errors)
                    errorList = errorList + "\nError: " + error;
            }
            /*else
            {
                fila++;
            }
        }
        catch (Exception ex)
        {                    
            _cells.GetCell(ResultColumn01, fila).Style = StyleConstants.Error;
            _cells.GetCell(ResultColumn01, fila).Value = "ERROR: " + ex.Message;
            Debugger.LogError("RibbonEllipse.cs:UpdateReferenceCodes()", ex.Message);
        }
        finally
        {
            fila++;
        }
    }*/
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
            if (_cells != null) _cells.SetCursorDefault();
        }

        private void drpEnviroment_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    _thread = new Thread(Reconsultar);

                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }

        private void btnAbout_Click_1(object sender, RibbonControlEventArgs e)
        {
            AboutBoxExcelAddIn About = new AboutBoxExcelAddIn("Gustavo Vargas", "");
            About.ShowDialog();
        }

        private void btnConsulta2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_excelApp.ActiveWorkbook.ActiveSheet.Name == SheetName01)
                {
                    //si ya hay un thread corriendo que no se ha detenido
                    if (_thread != null && _thread.IsAlive) return;
                    WoCode7 = "1";
                    _thread = new Thread(Consultar);
                    _thread.SetApartmentState(ApartmentState.STA);
                    _thread.Start();
                }
                else
                    MessageBox.Show(@"La hoja de Excel seleccionada no tiene el formato válido para realizar la acción");
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ReviewWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
        }
    }
}
