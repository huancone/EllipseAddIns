using System;
using System.Web.Services.Ellipse;
using System.Windows.Forms;
using SharedClassLibrary.Utilities;
using Authenticator = SharedClassLibrary.Ellipse.AuthenticatorService;

namespace SharedClassLibrary.Ellipse.Forms
{
    public partial class FormAuthenticate : Form
    {
        public string EllipseUser = "";
        public string EllipsePswd = "";
        public string EllipsePost = "";
        public string EllipseDstrct = "";

        // ReSharper disable once FieldCanBeMadeReadOny.Local
        //EllipseFunctions _eFunctions = new EllipseFunctions();
        public string SelectedEnvironment = null;

        public FormAuthenticate()
        {
            InitializeComponent();
            txtUsername.Text = EllipseUser;
            txtPassword.Text = EllipsePswd;
            drpPosition.Text = EllipsePost;
            txtDistrict.Text = EllipseDstrct;
        }
        public FormAuthenticate(string user, string password, string position, string district)
        {
            InitializeComponent();
            txtUsername.Text = EllipseUser = user;
            txtPassword.Text = EllipsePswd = password;
            drpPosition.Text = EllipsePost = position;
            txtDistrict.Text = EllipseDstrct = district;
        }
        public void SetUserAuthentication(string user, string password, string position, string district)
        {
            txtUsername.Text = EllipseUser = user;
            txtPassword.Text = EllipsePswd = password;
            drpPosition.Text = EllipsePost = position;
            txtDistrict.Text = EllipseDstrct = district;
        }
        private void btnAuthenticate_Click(object sender, EventArgs e)
        {
            EllipseUser = txtUsername.Text.ToUpper();
            EllipsePswd = txtPassword.Text;
            EllipsePost = (drpPosition.Text.Contains(" - ") ? drpPosition.Text.Substring(0, drpPosition.Text.IndexOf(" - ", StringComparison.Ordinal)).ToUpper() : EllipsePost = drpPosition.Text.ToUpper());
            EllipseDstrct = txtDistrict.Text.ToUpper();

            var authSer = new Authenticator.AuthenticatorService();
            var opAuth = new Authenticator.OperationContext
            {
                district = EllipseDstrct,
                position = EllipsePost,
                returnWarnings = true,
                returnWarningsSpecified = true
            };
            try
            {
                //control de selección de entorno en programación
                if (SelectedEnvironment == null)
                {
                    MessageBox.Show("Debe seleccionar un entorno de la lista para poder realizar la acción");
                    return;
                }

                authSer.Url = Connections.Environments.GetServiceUrl(SelectedEnvironment) + "/AuthenticatorService";
                ClientConversation.authenticate(EllipseUser, EllipsePswd, EllipseDstrct, EllipsePost);
                authSer.authenticate(opAuth);

                #region Form Population
                var positions = authSer.getPositions(opAuth);
                drpPosition.Items.Clear();
                foreach (var nvp in positions)
                    drpPosition.Items.Add(nvp.name + " - " + nvp.value);

                if (string.IsNullOrWhiteSpace(EllipsePost))
                    EllipsePost = positions[0].name.ToUpper();

                drpPosition.SelectedIndex = drpPosition.FindString(EllipsePost);

                if (string.IsNullOrWhiteSpace(txtDistrict.Text))
                {
                    var districts = authSer.getDistricts(opAuth);
                    if (districts != null && districts.Length > 0)
                        txtDistrict.Text = EllipseDstrct = districts[0].name.ToUpper();
                }
                else
                    txtDistrict.Text = txtDistrict.Text.ToUpper();

                #endregion
                DialogResult = DialogResult.OK;
                txtPassword.Text = "";
                //clearForm();
                Close();
            }
            catch(Exception ex)
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
                            txtDistrict.Text = EllipseDstrct = districts[0].name.ToUpper();
                    }
                    else
                        txtDistrict.Text = txtDistrict.Text.ToUpper();
                }
                catch(Exception exx)
                {
                    Debugger.LogError("FormAuthenticate:btnAuthenticate_Click(object, EventArgs):catch(catch)", exx.Message);
                }
                finally
                {
                    MessageBox.Show(@"Se ha producido un error al intentar realizar la autenticación. Asegúrese que los datos ingresados sean correctos e intente nuevamente." + "\n\n" + ex.Message);
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
