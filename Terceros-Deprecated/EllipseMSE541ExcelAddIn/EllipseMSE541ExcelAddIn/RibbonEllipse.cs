using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseMSE541ExcelAddIn.RefCodesService;
using EllipseMSE541ExcelAddIn.WorkRequestService;
using EllipseStdTextClassLibrary;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Excel.Application;
using OperationContext = EllipseMSE541ExcelAddIn.WorkRequestService.OperationContext;
using Screen = EllipseCommonsClassLibrary.ScreenService;
using Worksheet = Microsoft.Office.Tools.Excel.Worksheet;

namespace EllipseMSE541ExcelAddIn
{
    public partial class RibbonEllipse
    {
        ExcelStyleCells _cells;
        EllipseFunctions _eFunctions = new EllipseFunctions();
        FormAuthenticate _frmAuth = new FormAuthenticate();
        Application _excelApp;
        String _sheetName01 = "MSE541";
        Boolean _debugErrors = false;
        Worksheet _worksheet;
        String _colHeader = "AM";
        String _colFinal = "AN";
        String _colOcultar = "AO1";
        int _rowCabezera = 10;
        int _rowInicial = 11;
        int _maxRow = 10000;
        

        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            _excelApp = Globals.ThisAddIn.Application;
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);
            _eFunctions.DebugQueries = false;
            _eFunctions.DebugErrors = false;
            _eFunctions.DebugWarnings = false;

        }

        private void Ejecutar_Click(object sender, RibbonControlEventArgs e)
        {            
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

            if (_frmAuth.ShowDialog() == DialogResult.OK)
            //if(true)
            {
               //frmAuth.EllipseDstrct = "ICOR";
               //frmAuth.EllipsePost = "ADMIN";
               //frmAuth.EllipseUser = "hvilla1";
               //frmAuth.EllipsePswd = "h,1990";
               // Cells.getCell("A1").Value = "Conectado";

                var proxySheet = new WorkRequestService.WorkRequestService();
                var opSheet = new OperationContext();

                var proxySheetSla = new WorkRequestService.WorkRequestService();
                var opSheetSla = new OperationContext();

                var proxySheet2 = new RefCodesService.RefCodesService();
                var opSheet2 = new RefCodesService.OperationContext(); 

                var currentRow = _rowInicial;

                var requestId = "" + _cells.GetCell("A" + currentRow).Value;
                var status = "" + _cells.GetCell("B" + currentRow).Value;
                var desc1 = "" + _cells.GetCell("C" + currentRow).Value;
                var desc2 = "" + _cells.GetCell("D" + currentRow).Value;
                var classif = "" + _cells.GetCell("E" + currentRow).Value;
                var type = "" + _cells.GetCell("F" + currentRow).Value;
                var userStatus = "" + _cells.GetCell("G" + currentRow).Value;
                var prio = "" + _cells.GetCell("H" + currentRow).Value;
                var region = "" + _cells.GetCell("I" + currentRow).Value;
                var contactId = "" + _cells.GetCell("J" + currentRow).Value;
                var tySource = "" + _cells.GetCell("K" + currentRow).Value;
                var reference = "" + _cells.GetCell("L" + currentRow).Value;
                var equipo = "" + _cells.GetCell("M" + currentRow).Value;
                var empl = "" + _cells.GetCell("N" + currentRow).Value;
                var reqDate = "" + _cells.GetCell("O" + currentRow).Value;
                var reqTime = "" + _cells.GetCell("P" + currentRow).Value;
                var assig = "" + _cells.GetCell("Q" + currentRow).Value;
                var own = "" + _cells.GetCell("R" + currentRow).Value;
                var estimateNo = "" + _cells.GetCell("S" + currentRow).Value;
                var stdJob = "" + _cells.GetCell("T" + currentRow).Value;
                var stdJobDst = "" + _cells.GetCell("U" + currentRow).Value;
                var grupo = "" + _cells.GetCell("V" + currentRow).Value;
                var sla = "" + _cells.GetCell("W" + currentRow).Value;
                var fail = "" + _cells.GetCell("X" + currentRow).Value;
                var dueH = Convert.ToDecimal(_cells.GetCell("Y" + currentRow).Value);
                var descExt = "" + _cells.GetCell("Z" + currentRow).Value;
                var stoCode1 = "" + _cells.GetCell("AA" + currentRow).Value;
                var cant1 = "" + _cells.GetCell("AB" + currentRow).Value;
                var stoCode2 = "" + _cells.GetCell("AC" + currentRow).Value;
                var cant2 = "" + _cells.GetCell("AD" + currentRow).Value;
                var stoCode3 = "" + _cells.GetCell("AE" + currentRow).Value;
                var cant3 = "" + _cells.GetCell("AF" + currentRow).Value;
                var stoCode4 = "" + _cells.GetCell("AG" + currentRow).Value;
                var cant4 = "" + _cells.GetCell("AH" + currentRow).Value;
                var stoCode5 = "" + _cells.GetCell("AI" + currentRow).Value;
                var cant5 = "" + _cells.GetCell("AJ" + currentRow).Value;
                var otOrigen = "" + _cells.GetCell("AK" + currentRow).Value;    

                while (!string.IsNullOrEmpty(desc1))
                {                    
                        try
                        {
                            var requestParamsSheet = new WorkRequestServiceCreateRequestDTO();
                            var replySheet = new WorkRequestServiceCreateReplyDTO();                            
                            
                            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/WorkRequestService";

                            opSheet.district = _frmAuth.EllipseDstrct;
                            opSheet.position = _frmAuth.EllipsePost;
                            opSheet.maxInstances = 100;
                            opSheet.returnWarnings = _eFunctions.DebugWarnings;

                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            //requestParamsSheet.requestId = Cells.getNullIfTrimmedEmpty(request_id);
                            requestParamsSheet.requestIdDescription1 = desc1;
                            requestParamsSheet.requestIdDescription2 = desc2;
                            requestParamsSheet.classification = classif;
                            requestParamsSheet.requestType = type;
                            requestParamsSheet.userStatus = userStatus;
                            requestParamsSheet.priorityCode = prio;
                            requestParamsSheet.region = region;
                            requestParamsSheet.contactId = contactId;
                            requestParamsSheet.source = tySource;
                            requestParamsSheet.sourceReference = reference;
                            requestParamsSheet.equipmentRef = equipo;
                            requestParamsSheet.employee = empl;
                            requestParamsSheet.requiredByDate = reqDate;
                            requestParamsSheet.requiredByTime = reqTime;
                            requestParamsSheet.assignPerson = assig;                                
                            requestParamsSheet.ownerId = own;
                            requestParamsSheet.estimateNo = estimateNo;
                            requestParamsSheet.standardJob = stdJob;
                            requestParamsSheet.standardJobDistrict = stdJobDst;
                            requestParamsSheet.workGroup = grupo;
                                                    

                            replySheet = proxySheet.create(opSheet, requestParamsSheet);
                            requestId = replySheet.requestId;
                            status = replySheet.status;                                                        
                            
                            if (!string.IsNullOrEmpty(descExt))
                            {                                
                                StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "WQ" + replySheet.requestId, descExt);
                            }

                            if (!string.IsNullOrEmpty(stoCode1) || !string.IsNullOrEmpty(stoCode2) || !string.IsNullOrEmpty(stoCode3) || !string.IsNullOrEmpty(stoCode4) || !string.IsNullOrEmpty(stoCode5))
                            {
                                var requestParamsSheet2 = new RefCodesServiceModifyRequestDTO();
                                RefCodesServiceModifyReplyDTO replySheet2;

                                proxySheet2.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/RefCodesService";

                                opSheet2.district = _frmAuth.EllipseDstrct;
                                opSheet2.position = _frmAuth.EllipsePost;
                                opSheet2.maxInstances = 100;
                                opSheet2.returnWarnings = _eFunctions.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                if (!string.IsNullOrEmpty(stoCode1))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "001";
                                    requestParamsSheet2.refCode = stoCode1;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant1))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant1);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode2))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "002";
                                    requestParamsSheet2.refCode = stoCode2;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant2))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant2);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode3))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "003";
                                    requestParamsSheet2.refCode = stoCode3;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant3))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant3);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode4))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "004";
                                    requestParamsSheet2.refCode = stoCode4;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant4))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant4);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode5))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "005";
                                    requestParamsSheet2.refCode = stoCode5;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant5))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant5);
                                    }
                                }

                                if (!string.IsNullOrEmpty(otOrigen))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "009";
                                    requestParamsSheet2.seqNum = "001";
                                    requestParamsSheet2.refCode = otOrigen;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);                                    
                                }
                            }

                                if (!string.IsNullOrEmpty(sla))
                                {
                                    var requestParamsSheetSla = new WorkRequestServiceSetSLARequestDTO();
                                    WorkRequestServiceSetSLAReplyDTO replySheetSla;

                                    proxySheetSla.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/WorkRequestService";

                                    opSheetSla.district = _frmAuth.EllipseDstrct;
                                    opSheetSla.position = _frmAuth.EllipsePost;
                                    opSheetSla.maxInstances = 100;
                                    opSheetSla.returnWarnings = _eFunctions.DebugWarnings;

                                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                    requestParamsSheetSla.SLA = sla;
                                    requestParamsSheetSla.SLAFailureCode = fail;
                                    requestParamsSheetSla.requestId = requestId;
                                    requestParamsSheetSla.SLADueHours = dueH;

                                    replySheetSla = proxySheetSla.setSLA(opSheetSla,requestParamsSheetSla);
                                }                            

                            _cells.GetCell(1,currentRow).Value = requestId;
                            _cells.GetCell(2, currentRow).Value = status;
                            _cells.GetCell(_colFinal + currentRow).Value = "OK";
                            _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                            _cells.GetCell(_colFinal + currentRow).Select();
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(_colFinal + currentRow).Value = ex.Message;
                            _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                            _cells.GetCell(_colFinal + currentRow).Select();
                            Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);

                        }
                        finally
                        {
                            currentRow++;
                            desc1 = "" + _cells.GetCell("C" + currentRow).Value;
                            desc2 = "" + _cells.GetCell("D" + currentRow).Value;
                            classif = "" + _cells.GetCell("E" + currentRow).Value;
                            type = "" + _cells.GetCell("F" + currentRow).Value;
                            userStatus = "" + _cells.GetCell("G" + currentRow).Value;
                            prio = "" + _cells.GetCell("H" + currentRow).Value;
                            region = "" + _cells.GetCell("I" + currentRow).Value;
                            contactId = "" + _cells.GetCell("J" + currentRow).Value;
                            tySource = "" + _cells.GetCell("K" + currentRow).Value;
                            reference = "" + _cells.GetCell("L" + currentRow).Value;
                            equipo = "" + _cells.GetCell("M" + currentRow).Value;
                            empl = "" + _cells.GetCell("N" + currentRow).Value;
                            reqDate = "" + _cells.GetCell("O" + currentRow).Value;
                            reqTime = "" + _cells.GetCell("P" + currentRow).Value;
                            assig = "" + _cells.GetCell("Q" + currentRow).Value;
                            own = "" + _cells.GetCell("R" + currentRow).Value;
                            estimateNo = "" + _cells.GetCell("S" + currentRow).Value;
                            stdJob = "" + _cells.GetCell("T" + currentRow).Value;
                            stdJobDst = "" + _cells.GetCell("U" + currentRow).Value;
                            grupo = "" + _cells.GetCell("V" + currentRow).Value;
                            sla = "" + _cells.GetCell("W" + currentRow).Value;
                            fail = "" + _cells.GetCell("X" + currentRow).Value;
                            dueH = Convert.ToDecimal(_cells.GetCell("Y" + currentRow).Value);
                            descExt = "" + _cells.GetCell("Z" + currentRow).Value;
                            stoCode1 = "" + _cells.GetCell("AA" + currentRow).Value;
                            cant1 = "" + _cells.GetCell("AB" + currentRow).Value;
                            stoCode2 = "" + _cells.GetCell("AC" + currentRow).Value;
                            cant2 = "" + _cells.GetCell("AD" + currentRow).Value;
                            stoCode3 = "" + _cells.GetCell("AE" + currentRow).Value;
                            cant3 = "" + _cells.GetCell("AF" + currentRow).Value;
                            stoCode4 = "" + _cells.GetCell("AG" + currentRow).Value;
                            cant4 = "" + _cells.GetCell("AH" + currentRow).Value;
                            stoCode5 = "" + _cells.GetCell("AI" + currentRow).Value;
                            cant5 = "" + _cells.GetCell("AJ" + currentRow).Value;
                            otOrigen = "" + _cells.GetCell("AK" + currentRow).Value;
                        }                    
                }
                MessageBox.Show(@"Proceso Finalizado Correctamente");
            }
        }

        private void Formatear_Click(object sender, RibbonControlEventArgs e)
        {
            SetSheetHeaderData();
            Centrar();

            _worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            NamedRange groupRange;
            var groupCells = _worksheet.Range["A" + _rowInicial + ":"+ _colHeader + _maxRow];
            groupRange = _worksheet.Controls.AddNamedRange(groupCells, "GroupRange");

            groupRange.Change += AutoAjuste;

        }

        public void AutoAjuste(Range target)
        {
            _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
        }        

        public void SetSheetHeaderData()
        {
            try
            {
                _excelApp = Globals.ThisAddIn.Application;
                _excelApp.Workbooks.Add();
                _excelApp.ActiveWorkbook.ActiveSheet.Name = _sheetName01;

                if (_cells == null)

                    _cells = new ExcelStyleCells(_excelApp);

                _cells.GetCell(_colFinal + "1").Value = "OBLIGATORIO";
                _cells.GetCell(_colFinal + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_colFinal + "1").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.GetCell(_colFinal + "2").Value = "OPCIONAL";
                _cells.GetCell(_colFinal + "2").Style = _cells.GetStyle(StyleConstants.TitleOptional);
                _cells.GetCell(_colFinal + "3").Value = "INFORMATIVO";
                _cells.GetCell(_colFinal + "3").Style = _cells.GetStyle(StyleConstants.TitleInformation);
                _cells.GetCell(_colFinal + "4").Value = "ACCIÓN A REALIZAR";
                _cells.GetCell(_colFinal + "4").Style = _cells.GetStyle(StyleConstants.TitleAction);
                _cells.GetCell(_colFinal + "5").Value = "REQUERIDO ADICIONAL";
                _cells.GetCell(_colFinal + "5").Style = _cells.GetStyle(StyleConstants.TitleAdditional);

                _cells.GetRange(_colOcultar, "XFD1048576").Columns.Hidden = true;

                _cells.GetCell("A" + _rowCabezera).Value = "REQUEST_ID";
                _cells.GetCell("A" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("B" + _rowCabezera).Value = "STATUS";
                _cells.GetCell("B" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("C" + _rowCabezera).Value = "DESCRIPTION_1";
                _cells.GetCell("C" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("D" + _rowCabezera).Value = "DESCRIPTION_2";
                _cells.GetCell("D" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("E" + _rowCabezera).Value = "CLASSIF";
                _cells.GetCell("E" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleRequired);

                _cells.GetCell("F" + _rowCabezera).Value = "TYPE";
                _cells.GetCell("F" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("G" + _rowCabezera).Value = "USER_STATUS";
                _cells.GetCell("G" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("H" + _rowCabezera).Value = "PRIORITY";
                _cells.GetCell("H" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("I" + _rowCabezera).Value = "REGION";
                _cells.GetCell("I" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("J" + _rowCabezera).Value = "CONTACT_ID";
                _cells.GetCell("J" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("K" + _rowCabezera).Value = "TYPE_SOURCE";
                _cells.GetCell("K" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("L" + _rowCabezera).Value = "REFERENCE";
                _cells.GetCell("L" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("M" + _rowCabezera).Value = "EQUIP_NO";
                _cells.GetCell("M" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("N" + _rowCabezera).Value = "EMPLOYEED";
                _cells.GetCell("N" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("O" + _rowCabezera).Value = "REQ_DATE";
                _cells.GetCell("O" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("P" + _rowCabezera).Value = "REQ_TIME";
                _cells.GetCell("P" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("Q" + _rowCabezera).Value = "ASSIG_TO";
                _cells.GetCell("Q" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("R" + _rowCabezera).Value = "OWNER";
                _cells.GetCell("R" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("S" + _rowCabezera).Value = "ESTIMATE_NO";
                _cells.GetCell("S" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("T" + _rowCabezera).Value = "STD_JOB";
                _cells.GetCell("T" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("U" + _rowCabezera).Value = "STD_JOB_DIST";
                _cells.GetCell("U" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("V" + _rowCabezera).Value = "WORK_GROUP";
                _cells.GetCell("V" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);
                
                _cells.GetCell("W" + _rowCabezera).Value = "SLA";
                _cells.GetCell("W" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("X" + _rowCabezera).Value = "FAILURE_CODE";
                _cells.GetCell("X" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("Y" + _rowCabezera).Value = "DUE_HOURS";
                _cells.GetCell("Y" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("Z" + _rowCabezera).Value = "DESCRIPTION_EXTEND";
                _cells.GetCell("Z" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AA" + _rowCabezera).Value = "STOCK_CODE_1";
                _cells.GetCell("AA" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AB" + _rowCabezera).Value = "CANTIDAD";
                _cells.GetCell("AB" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AC" + _rowCabezera).Value = "STOCK_CODE_2";
                _cells.GetCell("AC" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AD" + _rowCabezera).Value = "CANTIDAD";
                _cells.GetCell("AD" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AE" + _rowCabezera).Value = "STOCK_CODE_3";
                _cells.GetCell("AE" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AF" + _rowCabezera).Value = "CANTIDAD";
                _cells.GetCell("AF" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AG" + _rowCabezera).Value = "STOCK_CODE_4";
                _cells.GetCell("AG" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AH" + _rowCabezera).Value = "CANTIDAD";
                _cells.GetCell("AH" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AI" + _rowCabezera).Value = "STOCK_CODE_5";
                _cells.GetCell("AI" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AJ" + _rowCabezera).Value = "CANTIDAD";
                _cells.GetCell("AJ" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AK" + _rowCabezera).Value = "OT_ORIGEN";
                _cells.GetCell("AK" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AL" + _rowCabezera).Value = "COMPLETED_BY";
                _cells.GetCell("AL" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AM" + _rowCabezera).Value = "CLOSED_DATE";
                _cells.GetCell("AM" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleOptional);

                _cells.GetCell("AN" + _rowCabezera).Value = "RESULTADO";
                _cells.GetCell("AN" + _rowCabezera).Style = _cells.GetStyle(StyleConstants.TitleInformation);

                _cells.GetCell("A6").Value = "GRUPO";
                _cells.GetCell("A6").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells("B6", "C6");
                _cells.SetValidationList(_cells.GetCell("B6"), GetGrupos());
                _cells.GetCell("B6").HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _cells.GetCell("B6").Font.Bold = true;

                _cells.MergeCells("D" + (_rowCabezera - 2), _colFinal + (_rowCabezera - 2));
                _cells.GetRange("D" + (_rowCabezera - 2), _colFinal + (_rowCabezera - 2)).Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("D" + (_rowCabezera - 2), _colFinal + (_rowCabezera - 2)).Borders.Weight = "2";

                _cells.MergeCells("D" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 3));
                _cells.GetRange("D" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 3)).Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("D" + (_rowCabezera - 3), _colFinal + (_rowCabezera - 3)).Borders.Weight = "2";

                _cells.MergeCells("D" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4));
                _cells.GetRange("D" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4)).Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("D" + (_rowCabezera - 4), _colFinal + (_rowCabezera - 4)).Borders.Weight = "2";
 

                _cells.GetCell("A7").Value = "ESTADO";
                _cells.GetCell("A7").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells("B7", "C7");
                _cells.SetValidationList(_cells.GetCell("B7"), GetEstados());
                _cells.GetCell("B7").HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _cells.GetCell("B7").Font.Bold = true;

                _cells.GetCell("A8").Value = "FECHA >=";
                _cells.GetCell("A8").Style = _cells.GetStyle(StyleConstants.TitleRequired);
                _cells.MergeCells("B8", "C8");
                _cells.GetCell("B8").HorizontalAlignment = XlHAlign.xlHAlignCenter;
                _cells.GetCell("B8").Font.Bold = true;

                _cells.GetRange("A6", "C7").Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("A6", "C7").Borders.Weight = "2";
                
                _cells.GetCell("A1").Value = "CERREJÓN";
                _cells.GetCell("A1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("A1", "A5");
                _cells.GetRange("A1", "A5").Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("A1", "A5").Borders.Weight = "2";

                _cells.GetCell("B1").Value = "WORKREQUEST - ELLIPSE";
                _cells.GetCell("B1").Style = _cells.GetStyle(StyleConstants.HeaderDefault);
                _cells.MergeCells("B1", _colHeader + "5");
                _cells.GetRange("B1", _colHeader + "5").Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("B1", _colHeader + "5").Borders.Weight = "2";
                /*Cells.mergeCells("C6", "L11");
                Cells.getRange("C6", "L11").Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                Cells.getRange("C6", "L11").Borders.Weight = "2";
                
                */
                _cells.MergeCells("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1));
                _cells.GetRange("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1)).Borders.Color = ColorTranslator.ToOle(Color.Black);
                _cells.GetRange("A" + (_rowCabezera - 1), _colFinal + (_rowCabezera - 1)).Borders.Weight = "2";
                                
                _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();

                _cells.GetCell("A" + _rowInicial).Select();
            }
            catch (Exception ex)
            {
                Debugger.LogError("ExcelStyleCells:setSheetHeaderData()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _debugErrors);
                MessageBox.Show(ex.Message);
            }
        }

        private void Centrar()
        {
            var clasificacion = GetClasificacion();
            var tipos = GetTipo();
            var estadosUsuarios = GetestadoUsuario();
            var prioridades = GetPrioridades();
            var regiones = GetRegiones();
            var fuentes = GetFuentes();
            var sla = GetSla();
            var fallas = GetFallas();

                _cells.GetCell("B" + _rowInicial + ":" + _colHeader + _maxRow).HorizontalAlignment = XlHAlign.xlHAlignCenter; 
                _cells.GetCell("A" + _rowInicial + ":A" + _maxRow).NumberFormat = "@";
                _cells.SetValidationList(_cells.GetCell("E" + _rowInicial + ":E" + _maxRow), clasificacion);
                _cells.GetCell("F" + _rowInicial + ":F" + _maxRow).NumberFormat = "@";
                _cells.SetValidationList(_cells.GetCell("F" + _rowInicial + ":F" + _maxRow), tipos);
                _cells.SetValidationList(_cells.GetCell("G" + _rowInicial + ":G" + _maxRow), estadosUsuarios);
                _cells.GetCell("H" + _rowInicial + ":H" + _maxRow).NumberFormat = "@";
                _cells.SetValidationList(_cells.GetCell("H" + _rowInicial + ":H" + _maxRow), prioridades);
                _cells.SetValidationList(_cells.GetCell("I" + _rowInicial + ":I" + _maxRow), regiones);
                _cells.SetValidationList(_cells.GetCell("K" + _rowInicial + ":K" + _maxRow), fuentes);
                _cells.GetCell("M" + _rowInicial + ":M" + _maxRow).NumberFormat = "@";
                _cells.GetCell("P" + _rowInicial + ":P" + _maxRow).NumberFormat = "@";
                _cells.SetValidationList(_cells.GetCell("W" + _rowInicial + ":W" + _maxRow), sla);
                _cells.SetValidationList(_cells.GetCell("X" + _rowInicial + ":X" + _maxRow), fallas);
                _cells.GetCell("X" + _rowInicial + ":X" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AA" + _rowInicial + ":AA" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AB" + _rowInicial + ":AB" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AC" + _rowInicial + ":AC" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AD" + _rowInicial + ":AD" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AE" + _rowInicial + ":AE" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AF" + _rowInicial + ":AF" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AG" + _rowInicial + ":AG" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AH" + _rowInicial + ":AH" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AI" + _rowInicial + ":AI" + _maxRow).NumberFormat = "@";
                _cells.GetCell("AJ" + _rowInicial + ":AJ" + _maxRow).NumberFormat = "@";

        }

        private void Centrar_Resultado(int i)
        {
            _cells.GetCell("B" + i + ":AK" + i).HorizontalAlignment = XlHAlign.xlHAlignCenter;                          
        }

        private void Limpiar()
        {
                _cells.GetCell("A" + _rowInicial + ":" + _colHeader + _maxRow).ClearContents();
                _cells.GetCell(_colFinal + _rowInicial + ":" + _colFinal + _maxRow).Clear();       
        }

        private void Limpiar_Busqueda(int i)
        {
            _cells.GetCell("B" + i + ":" + _colHeader + i).ClearContents();                
        }

       

        public List<string> GetClasificacion()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'RQCL'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getClasificacion = new List<string>();

            while (odr.Read())
            {
                getClasificacion.Add("" + odr["table_code"]);
            }
            return getClasificacion;
        }

        public List<string> GetGrupos()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select work_group from ellipse.msf720 where WG_STATUS = 'A'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getGrupos = new List<string>();

            while (odr.Read())
            {
                getGrupos.Add("" + odr["work_group"]);
            }
            return getGrupos;
        }

        public List<string> GetEstados()
        {

            var getEstados = new List<string>();

            getEstados.Add("C - CLOSED");
            getEstados.Add("O - OPEN");
            getEstados.Add("W - WORK");

            return getEstados;
        }
        
        public List<string> GetTipo()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'WO'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getTipo = new List<string>();

            while (odr.Read())
            {
                getTipo.Add("" + odr["table_code"]);
            }
            return getTipo;
        }

        public List<string> GetestadoUsuario()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'WS'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getestadoUsuario = new List<string>();

            while (odr.Read())
            {
                getestadoUsuario.Add("" + odr["table_code"]);
            }
            return getestadoUsuario;
        }

        public List<string> GetPrioridades()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'PY'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getPrioridades = new List<string>();

            while (odr.Read())
            {
                getPrioridades.Add("" + odr["table_code"]);
            }
            return getPrioridades;
        }

        public List<string> GetRegiones()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'REGN'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getRegiones = new List<string>();

            while (odr.Read())
            {
                getRegiones.Add("" + odr["table_code"]);
            }
            return getRegiones;
        }

        public List<string> GetFuentes()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'RQSC'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getFuentes = new List<string>();

            while (odr.Read())
            {
                getFuentes.Add("" + odr["table_code"]);
            }
            return getFuentes;
        }

        public List<string> GetSla()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'RQSL'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getSla = new List<string>();

            while (odr.Read())
            {
                getSla.Add("" + odr["table_code"]);
            }
            return getSla;
        }

        public List<string> GetFallas()
        {
            _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);

            var sqlQuery = "select trim(table_code) as table_code from ellipse.msf010 WHERE table_type = 'RQFC'";

            var odr = _eFunctions.GetQueryResult(sqlQuery);

            var getFallas = new List<string>();

            while (odr.Read())
            {
                getFallas.Add("" + odr["table_code"]);
            }
            return getFallas;
        }

        public bool IsNumeric(object expression)
        {

            bool isNum;

            double retNum;

            isNum = Double.TryParse(Convert.ToString(expression), NumberStyles.Any, NumberFormatInfo.InvariantInfo, out retNum);

            return isNum;

        }

        public void Consultar()
        {
            var i = _rowInicial;
            var wg = _cells.GetEmptyIfNull(_cells.GetCell("B6").Value);
            var es = _cells.GetEmptyIfNull(_cells.GetCell("B7").Value);
            var fecha = "" + _cells.GetCell("B8").Value;
            es = es.Substring(0, 1);
            
           /* if(IsNumeric(wq))
            {
                wq = wq.PadLeft(12, '0');
            }*/            

            while (!string.IsNullOrEmpty(wg))
            {
                try
                {
                    _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);                    

                    var sqlQuery = "SELECT " +
                                         "w.request_id, " +
                                         "w.request_stat, " +
                                         "w.short_desc_1, " +
                                         "w.short_desc_2, " +
                                         "w.work_req_classif, " +
                                         "w.work_req_type, " +
                                         "w.request_ustat, " +
                                         "w.priority_cde_541, " +
                                         "w.region, " +
                                         "w.contact_id, " +
                                         "w.work_req_source, " +
                                         "w.source_ref, " +
                                         "w.equip_no, " +
                                         "w.employee_id, " +
                                         "w.required_date, " +
                                         "w.required_time, " +
                                         "w.assign_person, " +
                                         "w.owner_id, " +
                                         "w.estimate_no, " +
                                         "w.std_job_no, " +
                                         "w.std_job_dstrct, " +
                                         "w.work_group, " +
                                         "w.sl_agreement, " +
                                         "w.sla_failure_code, " +
                                         "w.sla_due_hours, " +
                                         "w.completed_by, " +
                                         "w.closed_date, " +
                                         "trim(ts.std_volat_1||ts.std_volat_2||ts.std_volat_3||ts.std_volat_4||ts.std_volat_5) as desc_ext, " +
                                         "(select trim(ref_code) from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '001' and entity_value = w.request_id) as stock1, " +
                                         "(select trim(std_volat_1||std_volat_2||std_volat_3||std_volat_4||std_volat_5) from ellipse.msf096_std_volat where std_text_code = 'RC' and std_key = (select std_txt_key from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '001' and entity_value = w.request_id)) as cant1, " +
                                         "(select trim(ref_code) from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '002' and entity_value = w.request_id) as stock2, " +
                                         "(select trim(std_volat_1||std_volat_2||std_volat_3||std_volat_4||std_volat_5) from ellipse.msf096_std_volat where std_text_code = 'RC' and std_key = (select std_txt_key from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '002' and entity_value = w.request_id)) as cant2, " +
                                         "(select trim(ref_code) from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '003' and entity_value = w.request_id) as stock3, " +
                                         "(select trim(std_volat_1||std_volat_2||std_volat_3||std_volat_4||std_volat_5) from ellipse.msf096_std_volat where std_text_code = 'RC' and std_key = (select std_txt_key from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '003' and entity_value = w.request_id)) as cant3, " +
                                         "(select trim(ref_code) from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '004' and entity_value = w.request_id) as stock4, " +
                                         "(select trim(std_volat_1||std_volat_2||std_volat_3||std_volat_4||std_volat_5) from ellipse.msf096_std_volat where std_text_code = 'RC' and std_key = (select std_txt_key from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '004' and entity_value = w.request_id)) as cant4, " +
                                         "(select trim(ref_code) from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '005' and entity_value = w.request_id) as stock5, " +
                                         "(select trim(std_volat_1||std_volat_2||std_volat_3||std_volat_4||std_volat_5) from ellipse.msf096_std_volat where std_text_code = 'RC' and std_key = (select std_txt_key from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '001' and seq_num = '005' and entity_value = w.request_id)) as cant5, " +
                                         "(select trim(ref_code) from ellipse.msf071 where entity_type = 'WRQ' and ref_no = '009' and seq_num = '001' and entity_value = w.request_id) as ot_origen " +
                                         "FROM " +
                                         "ellipse.msf541 w, " +
                                         "ellipse.msf096_std_volat ts " +
                                         "where " +
                                         "w.work_group = " + "'" + wg + "'" +
                                         "and w.request_stat = " + "'" + es + "'" +
                                         "and w.raised_date >= " + "'" + fecha + "'" +
                                         " and ts.std_key(+) = w.request_id";

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    Limpiar();

                    while (odr.Read())
                    {
                        _cells.GetCell("A" + i).Value = odr["request_id"] + "";
                        _cells.GetCell("B" + i).Value = odr["request_stat"] + "";
                        _cells.GetCell("C" + i).Value = odr["short_desc_1"] + "";
                        _cells.GetCell("D" + i).Value = odr["short_desc_2"] + "";
                        _cells.GetCell("E" + i).Value = odr["work_req_classif"] + "";
                        _cells.GetCell("F" + i).Value = odr["work_req_type"] + "";
                        _cells.GetCell("G" + i).Value = odr["request_ustat"] + "";
                        _cells.GetCell("H" + i).Value = odr["priority_cde_541"] + "";
                        _cells.GetCell("I" + i).Value = odr["region"] + "";
                        _cells.GetCell("J" + i).Value = odr["contact_id"] + "";
                        _cells.GetCell("K" + i).Value = odr["work_req_source"] + "";
                        _cells.GetCell("L" + i).Value = odr["source_ref"] + "";
                        _cells.GetCell("M" + i).Value = odr["equip_no"] + "";
                        _cells.GetCell("N" + i).Value = odr["employee_id"] + "";
                        _cells.GetCell("O" + i).Value = odr["required_date"] + "";
                        _cells.GetCell("P" + i).Value = odr["required_time"] + "";
                        _cells.GetCell("Q" + i).Value = odr["assign_person"] + "";
                        _cells.GetCell("R" + i).Value = odr["owner_id"] + "";
                        _cells.GetCell("S" + i).Value = odr["estimate_no"] + "";
                        _cells.GetCell("T" + i).Value = odr["std_job_no"] + "";
                        _cells.GetCell("U" + i).Value = odr["std_job_dstrct"] + "";
                        _cells.GetCell("V" + i).Value = odr["work_group"] + "";
                        _cells.GetCell("W" + i).Value = odr["sl_agreement"] + "";
                        _cells.GetCell("X" + i).Value = odr["sla_failure_code"] + "";
                        _cells.GetCell("Y" + i).Value = odr["sla_due_hours"] + "";
                        _cells.GetCell("Z" + i).Value = odr["desc_ext"] + "";
                        _cells.GetCell("AA" + i).Value = odr["stock1"] + "";
                        _cells.GetCell("AB" + i).Value = odr["cant1"] + "";
                        _cells.GetCell("AC" + i).Value = odr["stock2"] + "";
                        _cells.GetCell("AD" + i).Value = odr["cant2"] + "";
                        _cells.GetCell("AE" + i).Value = odr["stock3"] + "";
                        _cells.GetCell("AF" + i).Value = odr["cant3"] + "";
                        _cells.GetCell("AG" + i).Value = odr["stock4"] + "";
                        _cells.GetCell("AH" + i).Value = odr["cant4"] + "";
                        _cells.GetCell("AI" + i).Value = odr["stock5"] + "";
                        _cells.GetCell("AJ" + i).Value = odr["cant5"] + "";
                        _cells.GetCell("AK" + i).Value = odr["ot_origen"] + "";
                        _cells.GetCell("AL" + i).Value = odr["completed_by"] + "";
                        _cells.GetCell("AM" + i).Value = odr["closed_date"] + "";

                        i++;

                        _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                    }

                    Centrar_Resultado(i);

                }
                catch (NullReferenceException)
                {
                    _cells.GetCell(i + 1, i).Value = "No fue Posible Obtener Informacion!";
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);

                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }
                finally
                {
                    wg = "";
                }
            }
        }

        private void Consulta_Click(object sender, RibbonControlEventArgs e)
        {
            Consultar();
        }

        private void clean_Click(object sender, RibbonControlEventArgs e)
        {
            Limpiar();
        }

        private void modificar_Click_1(object sender, RibbonControlEventArgs e)
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

            if (_frmAuth.ShowDialog() == DialogResult.OK)
            //if(true)
            {
               // frmAuth.EllipseDstrct = "ICOR";
               // frmAuth.EllipsePost = "ADMIN";
               // frmAuth.EllipseUser = "hvilla1";
               // frmAuth.EllipsePswd = "h,1990";


               // Cells.getCell("A1").Value = "Conectado";

                var proxySheet = new WorkRequestService.WorkRequestService();
                var opSheet = new OperationContext();

                var proxySheetSla = new WorkRequestService.WorkRequestService();
                var opSheetSla = new OperationContext();

                var proxySheet2 = new RefCodesService.RefCodesService();
                var opSheet2 = new RefCodesService.OperationContext(); 

                var currentRow = _rowInicial;

                var requestId = "" + _cells.GetCell("A" + currentRow).Value;
                var status = "" + _cells.GetCell("B" + currentRow).Value;
                var desc1 = "" + _cells.GetCell("C" + currentRow).Value;
                var desc2 = "" + _cells.GetCell("D" + currentRow).Value;
                var classif = "" + _cells.GetCell("E" + currentRow).Value;
                var type = "" + _cells.GetCell("F" + currentRow).Value;
                var userStatus = "" + _cells.GetCell("G" + currentRow).Value;
                var prio = "" + _cells.GetCell("H" + currentRow).Value;
                var region = "" + _cells.GetCell("I" + currentRow).Value;
                var contactId = "" + _cells.GetCell("J" + currentRow).Value;
                var tySource = "" + _cells.GetCell("K" + currentRow).Value;
                var reference = "" + _cells.GetCell("L" + currentRow).Value;
                var equipo = "" + _cells.GetCell("M" + currentRow).Value;
                var empl = "" + _cells.GetCell("N" + currentRow).Value;
                var reqDate = "" + _cells.GetCell("O" + currentRow).Value;
                var reqTime = "" + _cells.GetCell("P" + currentRow).Value;
                var assig = "" + _cells.GetCell("Q" + currentRow).Value;
                var own = "" + _cells.GetCell("R" + currentRow).Value;
                var estimateNo = "" + _cells.GetCell("S" + currentRow).Value;
                var stdJob = "" + _cells.GetCell("T" + currentRow).Value;
                var stdJobDst = "" + _cells.GetCell("U" + currentRow).Value;
                var grupo = "" + _cells.GetCell("V" + currentRow).Value;
                var sla = "" + _cells.GetCell("W" + currentRow).Value;
                var fail = "" + _cells.GetCell("X" + currentRow).Value;
                var dueH = Convert.ToDecimal(_cells.GetCell("Y" + currentRow).Value);
                var descExt = "" + _cells.GetCell("Z" + currentRow).Value;
                var stoCode1 = "" + _cells.GetCell("AA" + currentRow).Value;
                var cant1 = "" + _cells.GetCell("AB" + currentRow).Value;
                var stoCode2 = "" + _cells.GetCell("AC" + currentRow).Value;
                var cant2 = "" + _cells.GetCell("AD" + currentRow).Value;
                var stoCode3 = "" + _cells.GetCell("AE" + currentRow).Value;
                var cant3 = "" + _cells.GetCell("AF" + currentRow).Value;
                var stoCode4 = "" + _cells.GetCell("AG" + currentRow).Value;
                var cant4 = "" + _cells.GetCell("AH" + currentRow).Value;
                var stoCode5 = "" + _cells.GetCell("AI" + currentRow).Value;
                var cant5 = "" + _cells.GetCell("AJ" + currentRow).Value;
                var otOrigen = "" + _cells.GetCell("AK" + currentRow).Value;    

                while (!string.IsNullOrEmpty(desc1))
                {                    
                        try
                        {
                            var requestParamsSheet = new WorkRequestServiceModifyRequestDTO();
                            var replySheet = new WorkRequestServiceModifyReplyDTO();                            
                            
                            proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/WorkRequestService";

                            opSheet.district = _frmAuth.EllipseDstrct;
                            opSheet.position = _frmAuth.EllipsePost;
                            opSheet.maxInstances = 100;
                            opSheet.returnWarnings = _eFunctions.DebugWarnings;

                            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                            requestParamsSheet.requestId = requestId;
                            requestParamsSheet.requestIdDescription1 = desc1;
                            requestParamsSheet.requestIdDescription2 = desc2;
                            requestParamsSheet.classification = classif;
                            requestParamsSheet.requestType = type;
                            requestParamsSheet.userStatus = userStatus;
                            requestParamsSheet.priorityCode = prio;
                            requestParamsSheet.region = region;
                            requestParamsSheet.contactId = contactId;
                            requestParamsSheet.source = tySource;
                            requestParamsSheet.sourceReference = reference;
                            requestParamsSheet.equipmentRef = equipo;
                            requestParamsSheet.employee = empl;
                            requestParamsSheet.requiredByDate = reqDate;
                            requestParamsSheet.requiredByTime = reqTime;
                            requestParamsSheet.assignPerson = assig;                                
                            requestParamsSheet.ownerId = own;
                            requestParamsSheet.estimateNo = estimateNo;
                            requestParamsSheet.standardJob = stdJob;
                            requestParamsSheet.standardJobDistrict = stdJobDst;
                            requestParamsSheet.workGroup = grupo;
                                                    

                            replySheet = proxySheet.modify(opSheet, requestParamsSheet);
                            requestId = replySheet.requestId;
                            status = replySheet.status;                                                        
                            
                            if (!string.IsNullOrEmpty(descExt))
                            {                                
                                StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "WQ" + replySheet.requestId, descExt);
                            }

                            if (!string.IsNullOrEmpty(stoCode1) || !string.IsNullOrEmpty(stoCode2) || !string.IsNullOrEmpty(stoCode3) || !string.IsNullOrEmpty(stoCode4) || !string.IsNullOrEmpty(stoCode5))
                            {
                                var requestParamsSheet2 = new RefCodesServiceModifyRequestDTO();
                                var replySheet2 = new RefCodesServiceModifyReplyDTO();

                                proxySheet2.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/RefCodesService";

                                opSheet2.district = _frmAuth.EllipseDstrct;
                                opSheet2.position = _frmAuth.EllipsePost;
                                opSheet2.maxInstances = 100;
                                opSheet2.returnWarnings = _eFunctions.DebugWarnings;

                                ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                if (!string.IsNullOrEmpty(stoCode1))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "001";
                                    requestParamsSheet2.refCode = stoCode1;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant1))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant1);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode2))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "002";
                                    requestParamsSheet2.refCode = stoCode2;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant2))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant2);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode3))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "003";
                                    requestParamsSheet2.refCode = stoCode3;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant3))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant3);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode4))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "004";
                                    requestParamsSheet2.refCode = stoCode4;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant4))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant4);
                                    }
                                }

                                if (!string.IsNullOrEmpty(stoCode5))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "001";
                                    requestParamsSheet2.seqNum = "005";
                                    requestParamsSheet2.refCode = stoCode5;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);

                                    if (!string.IsNullOrEmpty(cant5))
                                    {
                                        StdText.SetCustomText(_eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label), StdText.GetCustomOpContext(_frmAuth.EllipseDstrct, _frmAuth.EllipsePost, 100, false), "RC" + replySheet2.stdTxtKey, cant5);
                                    }
                                }

                                if (!string.IsNullOrEmpty(otOrigen))
                                {
                                    requestParamsSheet2.entityType = "WRQ";
                                    requestParamsSheet2.refNo = "009";
                                    requestParamsSheet2.seqNum = "001";
                                    requestParamsSheet2.refCode = otOrigen;
                                    requestParamsSheet2.entityValue = requestId;

                                    replySheet2 = proxySheet2.modify(opSheet2, requestParamsSheet2);                                    
                                }

                                if (!string.IsNullOrEmpty(sla))
                                {
                                    var requestParamsSheetSla = new WorkRequestServiceSetSLARequestDTO();
                                    var replySheetSla = new WorkRequestServiceSetSLAReplyDTO();

                                    proxySheetSla.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/WorkRequestService";

                                    opSheetSla.district = _frmAuth.EllipseDstrct;
                                    opSheetSla.position = _frmAuth.EllipsePost;
                                    opSheetSla.maxInstances = 100;
                                    opSheetSla.returnWarnings = _eFunctions.DebugWarnings;

                                    ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                                    requestParamsSheetSla.SLA = sla;
                                    requestParamsSheetSla.SLAFailureCode = fail;
                                    requestParamsSheetSla.requestId = requestId;
                                    requestParamsSheetSla.SLADueHours = dueH;

                                    replySheetSla = proxySheetSla.setSLA(opSheetSla,requestParamsSheetSla);
                                }
                            }

                            _cells.GetCell(1,currentRow).Value = requestId;
                            _cells.GetCell(2, currentRow).Value = status;
                            _cells.GetCell(_colFinal + currentRow).Value = "OK";
                            _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                            _cells.GetCell(_colFinal + currentRow).Select();
                        }
                        catch (Exception ex)
                        {
                            _cells.GetCell(_colFinal + currentRow).Value = ex.Message;
                            _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                            _cells.GetCell(_colFinal + currentRow).Select();
                            Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);

                        }
                        finally
                        {
                            currentRow++;
                            desc1 = "" + _cells.GetCell("C" + currentRow).Value;
                            desc2 = "" + _cells.GetCell("D" + currentRow).Value;
                            classif = "" + _cells.GetCell("E" + currentRow).Value;
                            type = "" + _cells.GetCell("F" + currentRow).Value;
                            userStatus = "" + _cells.GetCell("G" + currentRow).Value;
                            prio = "" + _cells.GetCell("H" + currentRow).Value;
                            region = "" + _cells.GetCell("I" + currentRow).Value;
                            contactId = "" + _cells.GetCell("J" + currentRow).Value;
                            tySource = "" + _cells.GetCell("K" + currentRow).Value;
                            reference = "" + _cells.GetCell("L" + currentRow).Value;
                            equipo = "" + _cells.GetCell("M" + currentRow).Value;
                            empl = "" + _cells.GetCell("N" + currentRow).Value;
                            reqDate = "" + _cells.GetCell("O" + currentRow).Value;
                            reqTime = "" + _cells.GetCell("P" + currentRow).Value;
                            assig = "" + _cells.GetCell("Q" + currentRow).Value;
                            own = "" + _cells.GetCell("R" + currentRow).Value;
                            estimateNo = "" + _cells.GetCell("S" + currentRow).Value;
                            stdJob = "" + _cells.GetCell("T" + currentRow).Value;
                            stdJobDst = "" + _cells.GetCell("U" + currentRow).Value;
                            grupo = "" + _cells.GetCell("V" + currentRow).Value;
                            sla = "" + _cells.GetCell("W" + currentRow).Value;
                            fail = "" + _cells.GetCell("X" + currentRow).Value;
                            dueH = Convert.ToDecimal(_cells.GetCell("Y" + currentRow).Value);
                            descExt = "" + _cells.GetCell("Z" + currentRow).Value;
                            stoCode1 = "" + _cells.GetCell("AA" + currentRow).Value;
                            cant1 = "" + _cells.GetCell("AB" + currentRow).Value;
                            stoCode2 = "" + _cells.GetCell("AC" + currentRow).Value;
                            cant2 = "" + _cells.GetCell("AD" + currentRow).Value;
                            stoCode3 = "" + _cells.GetCell("AE" + currentRow).Value;
                            cant3 = "" + _cells.GetCell("AF" + currentRow).Value;
                            stoCode4 = "" + _cells.GetCell("AG" + currentRow).Value;
                            cant4 = "" + _cells.GetCell("AH" + currentRow).Value;
                            stoCode5 = "" + _cells.GetCell("AI" + currentRow).Value;
                            cant5 = "" + _cells.GetCell("AJ" + currentRow).Value;
                            otOrigen = "" + _cells.GetCell("AK" + currentRow).Value;
                        }                    
                }
                MessageBox.Show("Proceso Finalizado Correctamente");
            }
        }

        private void cerrar_Click(object sender, RibbonControlEventArgs e)
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

            if (_frmAuth.ShowDialog() == DialogResult.OK)
            //if (true)
            {
               // frmAuth.EllipseDstrct = "ICOR";
               // frmAuth.EllipsePost = "ADMIN";
               // frmAuth.EllipseUser = "hvilla1";
               // frmAuth.EllipsePswd = "h,1990";


                // Cells.getCell("A1").Value = "Conectado";

                var proxySheet = new WorkRequestService.WorkRequestService();
                var opSheet = new OperationContext();

                var currentRow = _rowInicial;

                var requestId = "" + _cells.GetCell("A" + currentRow).Value;
                String status;
                var closedBy = "" + _cells.GetCell("AL" + currentRow).Value;
                var closedDate = "" + _cells.GetCell("AM" + currentRow).Value;

                while (!string.IsNullOrEmpty(requestId))
                {
                    try
                    {
                        var requestParamsSheet = new WorkRequestServiceCloseRequestDTO();                        
                        var replySheet = new WorkRequestServiceCloseReplyDTO();

                        proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/WorkRequestService";

                        opSheet.district = _frmAuth.EllipseDstrct;
                        opSheet.position = _frmAuth.EllipsePost;
                        opSheet.maxInstances = 100;
                        opSheet.returnWarnings = _eFunctions.DebugWarnings;

                        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                        requestParamsSheet.requestId = requestId;
                        requestParamsSheet.closedBy = closedBy;
                        requestParamsSheet.closedDate = closedDate;

                        replySheet = proxySheet.close(opSheet, requestParamsSheet);
                        status = replySheet.status;                      

                        _cells.GetCell(2, currentRow).Value = status;
                        _cells.GetCell(_colFinal + currentRow).Value = "OK";
                        _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                        _cells.GetCell(_colFinal + currentRow).Select();
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(_colFinal + currentRow).Value = ex.Message;
                        _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                        _cells.GetCell(_colFinal + currentRow).Select();
                        Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);

                    }
                    finally
                    {
                        currentRow++;
                        requestId = "" + _cells.GetCell("A" + currentRow).Value;
                    }
                }
                MessageBox.Show("Proceso Finalizado Correctamente");
            }
        }

        private void wrCerrar_Click(object sender, RibbonControlEventArgs e)
        {
                Limpiar();
                var i = _rowInicial;
            
                try
                {
                   
                    _eFunctions.SetDBSettings("SIGCOPRD", "sigman", "sig0679");

                    var sqlQuery = "SELECT " +
                                         "request_id, " +
                                         "completed_by, " +
                                         "closed_date " +
                                         "FROM " +
                                         "cierrawr order by 3 desc";

                    var odr = _eFunctions.GetQueryResult(sqlQuery);

                    while (odr.Read())
                    {
                        _cells.GetCell("A" + i).Value = odr["request_id"] + "";
                        _cells.GetCell("AL" + i).Value = odr["completed_by"] + "";
                        _cells.GetCell("AM" + i).Value = odr["closed_date"] + "";

                        i++;

                        _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                    }

                    Centrar_Resultado(i);

                }
                catch (NullReferenceException)
                {
                    _cells.GetCell(i + 1, i).Value = "No fue Posible Obtener Informacion!";
                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }
                catch (Exception error)
                {
                    var messageBox = MessageBox.Show(error.Message);

                    _excelApp.ActiveWorkbook.ActiveSheet.Cells.Columns.AutoFit();
                }

                //cerrar_automatico();
        }

        private void cerrar_automatico()
        {
            _frmAuth.StartPosition = FormStartPosition.CenterScreen;
            _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;

            //if (frmAuth.ShowDialog() == DialogResult.OK)
            if (true)
            {
                _frmAuth.EllipseDstrct = "ICOR";
                _frmAuth.EllipsePost = "AGSS";
                _frmAuth.EllipseUser = "interctd";
                _frmAuth.EllipsePswd = "fe0679";


                // Cells.getCell("A1").Value = "Conectado";

                var proxySheet = new WorkRequestService.WorkRequestService();
                var opSheet = new OperationContext();

                var currentRow = _rowInicial;

                var requestId = "" + _cells.GetCell("A" + currentRow).Value;
                String status;
                var closedBy = "" + _cells.GetCell("AL" + currentRow).Value;
                var closedDate = "" + _cells.GetCell("AM" + currentRow).Value;

                while (!string.IsNullOrEmpty(requestId))
                {
                    try
                    {
                        var requestParamsSheet = new WorkRequestServiceCloseRequestDTO();
                        var replySheet = new WorkRequestServiceCloseReplyDTO();

                        proxySheet.Url = _eFunctions.GetServicesUrl(drpEnvironment.SelectedItem.Label) + "/WorkRequestService";

                        opSheet.district = _frmAuth.EllipseDstrct;
                        opSheet.position = _frmAuth.EllipsePost;
                        opSheet.maxInstances = 100;
                        opSheet.returnWarnings = _eFunctions.DebugWarnings;

                        ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

                        requestParamsSheet.requestId = requestId;
                        requestParamsSheet.closedBy = closedBy;
                        requestParamsSheet.closedDate = closedDate;

                        replySheet = proxySheet.close(opSheet, requestParamsSheet);
                        status = replySheet.status;

                        _cells.GetCell(2, currentRow).Value = status;
                        _cells.GetCell(_colFinal + currentRow).Value = "OK";
                        _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAdditional);
                        _cells.GetCell(_colFinal + currentRow).Select();
                    }
                    catch (Exception ex)
                    {
                        _cells.GetCell(_colFinal + currentRow).Value = ex.Message;
                        _cells.GetCell(_colFinal + currentRow).Style = _cells.GetStyle(StyleConstants.TitleAction);
                        _cells.GetCell(_colFinal + currentRow).Select();
                        Debugger.LogError("RibbonEllipse:startLabourCostLoad()", "\n\rMessage:" + ex.Message + "\n\rSource:" + ex.Source + "\n\rStackTrace:" + ex.StackTrace, _eFunctions.DebugErrors);

                    }
                    finally
                    {
                        currentRow++;
                        requestId = "" + _cells.GetCell("A" + currentRow).Value;
                    }
                }
                MessageBox.Show("Proceso Finalizado Correctamente");
            }
        }

    }
}
