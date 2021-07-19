using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Vsto.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace EllipseTemplateExcelAddIn
{
    public partial class RibbonEllipse
    {
        private ExcelStyleCells _cells;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;
        private Excel.Application _excelApp;
        
        private Thread _thread;
        private void RibbonEllipse_Load(object sender, RibbonUIEventArgs e)
        {
            LoadSettings();
        }

        public void LoadSettings()
        {
            try
            {
                Settings.Initiate();
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

                var settings = Settings.CurrentSettings;
                settings.SetDefaultCustomSettingValue("ParameterBoolean", true);
                settings.SetDefaultCustomSettingValue("ParameterText", "value2");
                settings.SetDefaultCustomSettingValue("ParameterNumber", 12345);

                //Setting of Configuration Options from Config File (or default)
                try
                {
                    settings.LoadCustomSettings();
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                var varBool = MyUtilities.IsTrue(settings.GetCustomSettingValue("ParameterBoolean"));
                var varText = settings.GetCustomSettingValue("ParameterText");
                var varNumber = MyUtilities.ToInteger(settings.GetCustomSettingValue("ParameterNumber"));

                bool boolOption = varBool;
                string tbTextMessage = varText;
                double numberOption = varNumber * 0.3;
                //
                settings.SaveCustomSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

		private void btnExecution_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _frmAuth.StartPosition = FormStartPosition.CenterScreen;
                _frmAuth.SelectedEnvironment = drpEnvironment.SelectedItem.Label;
                if (_frmAuth.ShowDialog() != DialogResult.OK) return;

                //si ya hay un thread corriendo que no se ha detenido
                if (_thread != null && _thread.IsAlive) return;
                _thread = new Thread(ExecuteQuery);

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
            }
            catch (Exception ex)
            {
                Debugger.LogError("RibbonEllipse.cs:ExecuteQuery()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }

        }

        private void ExecuteQuery()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);
                _cells.SetCursorWait();

                var titleRow = 1;
                var sqlQuery = @"SELECT WORK_ORDER FROM ELLIPSE.MSF620 WO WHERE WO.RAISED_DATE = '20200801' AND WO.WORK_GROUP = 'MTOLOC'";
                var tableName = "table";
                _eFunctions.SetDBSettings(drpEnvironment.SelectedItem.Label);
                var dataReader = _eFunctions.GetQueryResult(sqlQuery);

                if (dataReader == null)
                    return;

                //Cargo el encabezado de la tabla y doy formato
                for (var i = 0; i < dataReader.FieldCount; i++)
                    _cells.GetCell(i + 1, titleRow).Value2 = "'" + dataReader.GetName(i);

                _cells.FormatAsTable(_cells.GetRange(1, titleRow, dataReader.FieldCount, titleRow + 1), tableName);

                //cargo los datos 
                if (dataReader.IsClosed) return;


                var currentRow = titleRow + 1;
                while (dataReader.Read())
                {
                    for (var i = 0; i < dataReader.FieldCount; i++)
                        _cells.GetCell(i + 1, currentRow).Value2 = "'" + dataReader[i].ToString().Trim();
                    currentRow++;
                }

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

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            new AboutBoxExcelAddIn("Desarrollador/Empresa").ShowDialog();
        }

        private void btnStopThread_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (_thread != null && _thread.IsAlive)
                    _thread.Abort();
                _cells?.SetCursorDefault();
            }
            catch (ThreadAbortException ex)
            {
                MessageBox.Show(@"Se ha detenido el proceso. " + ex.Message);
            }
        }
		private void GeneralService()
        {
            try
            {
                if (_cells == null)
                    _cells = new ExcelStyleCells(_excelApp);

                /*
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

                //Se cargan los parámetros de  la solicitud
                var request = new WorkOrderServiceCreateRequestDTO();
                request.districtCode = "ICOR";
                request.workGroup = "MTOLOC";
                request.workOrderDesc = "ORDEN DE PRUEBA";
                request.workOrderType = "CO";
                request.maintenanceType = "CO";
                request.equipmentNo = "1000016";

                //se envía la acción
                var reply = service.create(opContext, request);

                //se analiza la respuesta y se hacen las acciones pertinentes
                _cells.GetCell(1, 1).Value2 = reply.workOrder.prefix + reply.workOrder.no;
                */


            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Se ha producido un error al intentar crear la orden de trabajo." + "\n\n" + ex.Message);
            }
        }
        private void ScreenService()
        {
            //Proceso del Servicio Screen
            var service = new SharedClassLibrary.Ellipse.ScreenService.ScreenService();
            var urlService = Environments.GetServiceUrl(drpEnvironment.SelectedItem.Label);
            service.Url = urlService + "/ScreenService";

            //Instanciar el Contexto de Operación
            var opContext = new SharedClassLibrary.Ellipse.ScreenService.OperationContext
            {
                district = _frmAuth.EllipseDstrct,
                position = _frmAuth.EllipsePost
            };

            //Instanciar el SOAP
            ClientConversation.authenticate(_frmAuth.EllipseUser, _frmAuth.EllipsePswd);

            //Solicitud 1 y Respuesta 1
            var reply = service.executeScreen(opContext, "MSO435");

            //validamos el ingreso al programa
            if (reply.mapName != "MSM435A")
                throw new Exception("ERROR:" + "No se pudo establecer comunicación con el servicio");

            //arreglo para los campos del screen
            var arrayFields = new ArrayScreenNameValue();

            //se adicionan los campos que se vayan a enviar
            arrayFields.Add("OPTION1I", "1");
            arrayFields.Add("MODEL_CODE1I", "CÓDIGO MODELO");
            arrayFields.Add("STAT_DATE1I", "FECHA MODELO");
            arrayFields.Add("SHIFT1I", "TURNO MODELO");

            //Solicitud 2
            var request = new SharedClassLibrary.Ellipse.ScreenService.ScreenSubmitRequestDTO();
            request.screenFields = arrayFields.ToArray();
            request.screenKey = "1";

            //Respuesta 2
            reply = service.submit(opContext, request);

            //Existencia y nombre de pantalla de respuesta
            if (reply == null || reply.mapName == "MSM435B")
                throw new Exception("ERROR:" + "Se ha producido un error al intentar enviar la solicitud");

            //La respuesta tiene un error o una advertencia
            if (_eFunctions.CheckReplyError(reply) || _eFunctions.CheckReplyWarning(reply))
                throw new Exception("ERROR:" + "Se ha producido un error al intentar enviar la solicitud");

            //La respuesta pide confirmación
            if (reply.functionKeys == "XMIT-Confirm")
                reply = service.submit(opContext, request);

            //si necesitas obtener los campso del reply y trabajar con ellos
            var replyFields = new ArrayScreenNameValue(reply.screenFields);
            var woProject = replyFields.GetField("WO_PROJ1I1").value.Equals("");
        }

    }
}
