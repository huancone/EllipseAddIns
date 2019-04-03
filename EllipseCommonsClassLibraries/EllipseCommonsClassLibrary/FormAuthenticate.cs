using System;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using Authenticator = EllipseCommonsClassLibrary.AuthenticatorService;

namespace EllipseCommonsClassLibrary
{
    public partial class FormAuthenticate : Form
    {
        // ReSharper disable once FieldCanBeMadeReadOny.Local
        private readonly EllipseFunctions _eFunctions = new EllipseFunctions();
        public string EllipseDsct = "";
        public string EllipsePost = "";
        public string EllipsePswd = "";
        public string EllipseUser = "";
        public string SelectedEnviroment = null;

        public FormAuthenticate()
        {
            InitializeComponent();
            txtUsername.Text = EllipseUser;
            txtPassword.Text = EllipsePswd;
            txtDistrict.Text = EllipseDsct;
            drpPosition.Text = EllipsePost;
        }

        private void btnAuthenticate_Click(object sender, EventArgs e)
        {
            EllipseDsct = txtDistrict.Text.ToUpper();
            EllipsePost = drpPosition.Text.Contains(" - ")
                ? drpPosition.Text.Substring(0, drpPosition.Text.IndexOf(" - ", StringComparison.Ordinal)).ToUpper()
                : EllipsePost = drpPosition.Text.ToUpper();
            EllipsePswd = txtPassword.Text;
            EllipseUser = txtUsername.Text.ToUpper();

            var authSer = new Authenticator.AuthenticatorService();
            var opAuth = new Authenticator.OperationContext
            {
                district = EllipseDsct,
                position = EllipsePost,
                returnWarnings = true
            };

            //control de selección de entorno en programación
            if (SelectedEnviroment == null)
            {
                MessageBox.Show("Debe seleccionar un entorno de la lista para poder realizar la acción");
                return;
            }

            authSer.Url = _eFunctions.GetServicesUrl(SelectedEnviroment) + "/AuthenticatorService";
            try
            {
                ClientConversation.authenticate(EllipseUser, EllipsePswd, EllipseDsct, EllipsePost);
                authSer.authenticate(opAuth);

                #region Form Population

                var positions = authSer.getPositions(opAuth);
                drpPosition.Items.Clear();
                foreach (var nvp in positions)
                    drpPosition.Items.Add(nvp.name + " - " + nvp.value);

                if (string.IsNullOrWhiteSpace(EllipsePost))
                    EllipsePost = positions[0].name.ToUpper();

                drpPosition.SelectedIndex = drpPosition.FindString(EllipsePost);

                var districts = authSer.getDistricts(opAuth);
                if (districts != null && districts.Length > 0)
                    txtDistrict.Text = EllipseDsct = districts[0].name;

                #endregion

                DialogResult = DialogResult.OK;
                txtPassword.Text = "";
                //clearForm();
                Close();
            }
            catch (Exception ex)
            {
                try
                {
                    var positions = authSer.getPositions(opAuth);
                    drpPosition.Items.Clear();
                    foreach (var nvp in positions)
                        drpPosition.Items.Add(nvp.name + " - " + nvp.value);

                    var districts = authSer.getDistricts(opAuth);
                    if (districts != null && districts.Length > 0)
                        txtDistrict.Text = EllipseDsct = districts[0].name;
                }
                catch (Exception exx)
                {
                    Debugger.LogError("FormAuthenticate:btnAuthenticate_Click(object, EventArgs):catch(catch)",
                        exx.Message);
                }
                finally
                {
                    MessageBox.Show(
                        @"Se ha producido un error al intentar realizar la autenticación. Asegúrese que los datos ingresados sean correctos e intente nuevamente. \n\n" +
                        ex.Message);
                }
            }
        }

        public void ClearForm()
        {
            txtDistrict.Clear();
            drpPosition.Items.Clear();
            txtUsername.Clear();
            txtPassword.Clear();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}