using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using EllipseCommonsClassLibrary;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Classes;
using EllipseCommonsClassLibrary.Constants;
using EllipseCommonsClassLibrary.Utilities;
using System.Web.Services.Ellipse.Post;
using Microsoft.Office.Tools.Ribbon;
using TestAddinEllipse9.WorkOrderService;
using Application = Microsoft.Office.Interop.Excel.Application;
using FormAuthenticate = EllipseCommonsClassLibrary.FormAuthenticate;


namespace TestAddinEllipse9
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions = new EllipseFunctions();
        private FormAuthenticate _frmAuth = new FormAuthenticate();
        private Application _excelApp;

        private const string ValidationSheetName = "ValidationSheetWorkOrder";
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

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ClientConversation.debuggingMode = true;
                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK)
                    return;
                ConsultarOt();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:CreateWoList()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn().ShowDialog();
        }

        private void ConsultarOt()
        {
            if (_cells == null)
                _cells = new ExcelStyleCells(_excelApp);


            //Creación del Servicio
            var service = new WorkOrderService.WorkOrderService();
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            service.Url = urlService + "/WorkOrderService";

            //Instanciar el Contexto de Operación
            var opContext = new WorkOrderService.OperationContext
            {
                district = _frmAuth.EllipseDsct,
                position = _frmAuth.EllipsePost
            };

            //Instanciar el SOAP
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //Se cargan los parámetros de la solicitud
            string workOrder = "" + _cells.GetCell(1, 1).Value;
            _cells.GetCell(1, 2).Value2 = service.Url;

            var request = new WorkOrderServiceReadRequestDTO();
            request.districtCode = "ICOR";
            request.workOrder = new WorkOrderDTO();
            request.workOrder.prefix = workOrder.Substring(0, 2);
            request.workOrder.no = workOrder.Substring(2);
            request.includeTasks = false;
            request.includeTasksSpecified = true;

            //se envía la acción
            var reply = service.read(opContext, request);

            //se analiza la respuesta y se hacen las acciones pertinentes

            var i = 2;
            _cells.GetCell(1, i).Value2 = reply.workOrder.prefix + reply.workOrder.no;
            _cells.GetCell(2, i).Value2 = reply.workOrderDesc;
            _cells.GetCell(3, i).Value2 = reply.equipmentNo;
            _cells.GetCell(4, i).Value2 = reply.workOrderType;
            _cells.GetCell(5, i).Value2 = reply.maintenanceType;
            _cells.GetCell(6, i).Value2 = reply.raisedDate;
            _cells.GetCell(7, i).Value2 = reply.originatorId;



        }
    }
}
