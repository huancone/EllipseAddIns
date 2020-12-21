using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using LogsheetDatamodelLibrary.Configuration;
using SharedClassLibrary;
using SharedClassLibrary.Forms;
using SharedClassLibrary.Utilities;

namespace LogsheetDatamodelLibrary
{
    public class AuthenticationForm : SharedClassLibrary.Forms.FormAuthenticate
    {
        public override void CancelAction()
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        public override void AuthenticateAction()
        {
            try
            {
                LsdmConfig.Login.User = User = txtUsername.Text.ToUpper();
                LsdmConfig.Login.Password = Pswd = txtPassword.Text;

                DialogResult = DialogResult.OK;
                txtPassword.Text = "";
                //clearForm();
                Close();
            }
            catch (Exception ex)
            {
                try
                {

                }
                catch (Exception exx)
                {
                    Debugger.LogError("FormAuthenticate:btnAuthenticate_Click(object, EventArgs):catch(catch)", exx.Message);
                }
                finally
                {
                    MessageBox.Show(Resources.Autentication_Error + @". " + Resources.Error_ValidateInput + @"." + Environment.NewLine + Environment.NewLine + ex.Message);
                }
            }
        }
    }
}
