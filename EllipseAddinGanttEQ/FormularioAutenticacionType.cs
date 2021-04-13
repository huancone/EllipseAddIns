using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Authenticator = SharedClassLibrary.Ellipse.AuthenticatorService;
using System.Web.Services.Ellipse;
using SharedClassLibrary;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using Debugger = SharedClassLibrary.Utilities.Debugger;
using Oracle.ManagedDataAccess.Client;
//using Excel = Microsoft.Office.Interop.Excel;


namespace EllipseAddinGanttEQ
{
    public partial class FormularioAutenticacionType : Form
    {
        public string Permiso;
        public string EllipseUser = "";
        public string EllipsePswd = "";
        public string EllipsePost = "";
        public string EllipseDsct = "";

        private OracleConnection _sqlOracleConn;
        private int _connectionTimeOut = 30;//default ODP 15
        private bool _poolingDataBase = true;//default ODP true
        private OracleCommand _sqlOracleComm;

        EllipseFunctions _eFunctions;
        public string SelectedEnvironment = null;
        //private Excel.Application _excelApp;

        public FormularioAutenticacionType()
        {
            InitializeComponent();
            txtUsername.Text = EllipseUser;
            txtPassword.Text = EllipsePswd;
            txtDistrict.Text = EllipseDsct;
            drpPosition.Text = EllipsePost;
            //Permiso = txtUsername.Text;
        }

        private void btnAuthenticate_Click(object sender, EventArgs e)
        {

            //_excelApp.Visible = true;
            //_excelApp.ScreenUpdating = true;

            Authenticator.AuthenticatorService authSer = new Authenticator.AuthenticatorService();
            Authenticator.OperationContext opAuth = new Authenticator.OperationContext
            {
                district = EllipseDsct,
                position = EllipsePost,
                returnWarnings = true,
                returnWarningsSpecified = true
            };

            try
            {
                EllipseDsct = txtDistrict.Text.ToUpper();
                EllipsePost = (drpPosition.Text.Contains(" - ") ? drpPosition.Text.Substring(0, drpPosition.Text.IndexOf(" - ", StringComparison.Ordinal)).ToUpper() : EllipsePost = drpPosition.Text.ToUpper());
                EllipsePswd = txtPassword.Text;
                EllipseUser = txtUsername.Text.ToUpper();

                authSer.Url = Environments.GetServiceUrl("Productivo") + "/AuthenticatorService";
                ClientConversation.authenticate(EllipseUser, EllipsePswd);
                Authenticator.NameValuePair[] districts = authSer.getDistricts(opAuth);
                Authenticator.NameValuePair[] positionsx = authSer.getPositions(opAuth);
                EllipsePost = positionsx[0].name.ToUpper();
                EllipseDsct = districts[0].name.ToUpper();


                //control de selección de entorno en programación
                /*if (SelectedEnvironment == null)
                {
                    MessageBox.Show("Debe seleccionar un entorno de la lista para poder realizar la acción");
                    return;
                }
                */
                //Selecionar Item Posicion.

                authSer.authenticate(opAuth);




                if (_eFunctions == null)
                    _eFunctions = new EllipseFunctions();
                _eFunctions.SetDBSettings("sigcoprd", "sigman", "sig0679", "");

                    //var dataReader = _eFunctions.GetQueryResult(sqlQuery);
                    var dataReader = _eFunctions.GetQueryResult(@"SELECT
                                                           USUARIO  AS C_USER,
                                                           ROL
                                                        FROM
                                                            SIGMAN.USER_ADDIN
                                                        WHERE
                                                            USUARIO = '" + (txtUsername.Text).ToUpper() + @"'");


                //--AND PASSWD = '" + txtPassword.Text + "'"
                Int32 Contador = 0;
                while (dataReader.Read())
                {
                    Contador++;
                    Permiso = dataReader["ROL"].ToString();
                }

                if (Contador > 0)
                {
                    //MessageBox.Show("Bienvenido");
                    DialogResult = DialogResult.OK;
                    //txtPassword.Text = "";
                    //clearForm();
                    Close();

                    //return;
                }
                else
                {
                    MessageBox.Show("No tiene Acceso para aceder a esta aplicacion");
                    return;
                }


            }
            catch (Exception ex)
            {
                try
                {
                    var positions = authSer.getPositions(opAuth);
                    drpPosition.Items.Clear();
                    foreach (var nvp in positions)
                        drpPosition.Items.Add(nvp.name + " - " + nvp.value);

                    if (string.IsNullOrWhiteSpace(txtDistrict.Text))
                    {
                        var districts = authSer.getDistricts(opAuth);
                        if (districts != null && districts.Length > 0)
                            txtDistrict.Text = EllipseDsct = districts[0].name.ToUpper();
                    }
                    else
                        txtDistrict.Text = txtDistrict.Text.ToUpper();
                }
                catch (Exception exx)
                {
                    Debugger.LogError("FormAuthenticate:btnAuthenticate_Click(object, EventArgs):catch(catch)", exx.Message);
                }
                finally
                {
                    MessageBox.Show(@"Se ha producido un error al intentar realizar la autenticación. Asegúrese que los datos ingresados sean correctos e intente nuevamente." + "\n\n" + ex.Message);
                }
            }

        }

        private void Cancelar_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
