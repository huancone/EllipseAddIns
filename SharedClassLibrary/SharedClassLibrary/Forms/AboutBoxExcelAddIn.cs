using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using SharedClassLibrary.Configuration;

//Shared Class Library - AbourBoxExcelAddIn
//Desarrollado por:
//Héctor J Hernández R <hernandezrhectorj@gmail.com>
//Hugo A Mendoza B <hugo.mendoza@hambings.com.co>

// ReSharper disable ConvertPropertyToExpressionBody

namespace SharedClassLibrary.Forms
{
    public abstract partial class AboutBoxExcelAddIn : Form
    {
        private AssemblyItem addinAssembly;
        public AboutBoxExcelAddIn()
        {
            addinAssembly = new AssemblyItem(AssemblyItem.GetLastAssembly());
            InitializeComponent();
            Text = string.Format("About {0}", addinAssembly.AssemblyTitle);
            labelProductName.Text = addinAssembly.AssemblyProduct;
            labelVersion.Text = string.Format("Version {0}", addinAssembly.AssemblyVersion);
            labelCopyright.Text = addinAssembly.AssemblyCopyright;
            labelCompanyName.Text = addinAssembly.AssemblyCompany;
            textBoxDescription.Text = addinAssembly.AssemblyDescription;
            labelDeveloper1.Text = addinAssembly.AssemblyDeveloper1;
            labelDeveloper2.Text = addinAssembly.AssemblyDeveloper2;
            //var resourceFile = "CommonsClassLibrary.Resources.aboutPictureBox.png";
            //logoPictureBox.Image = Image.FromStream(assembly.GetManifestResourceStream(resourceFile));
        }

        public AboutBoxExcelAddIn(string developerName1, string developerName2)
        {
            addinAssembly = new AssemblyItem(AssemblyItem.GetLastAssembly());

            InitializeComponent();
            Text = string.Format("About {0}", addinAssembly.AssemblyTitle);
            labelProductName.Text = addinAssembly.AssemblyProduct;
            labelVersion.Text = string.Format("Version {0}", addinAssembly.AssemblyVersion);
            labelCopyright.Text = addinAssembly.AssemblyCopyright;
            labelCompanyName.Text = addinAssembly.AssemblyCompany;
            textBoxDescription.Text = addinAssembly.AssemblyDescription;
            labelDeveloper1.Text = developerName1;
            labelDeveloper2.Text = developerName2;
            //var resourceFile = "CommonsClassLibrary.Resources.aboutPictureBox.png";
            //logoPictureBox.Image = Image.FromStream(assembly.GetManifestResourceStream(resourceFile));
        }

        public void ShowAdvancedSettingsButton(bool status)
        {
            btnAdvancedSettings.Enabled = status;
            btnAdvancedSettings.Visible = status;
        }

        private void tableLayoutPanel_Paint(object sender, PaintEventArgs e)
        {
        }

        private void btnRepository_Click(object sender, EventArgs e)
        {
            OpenRepository();
        }

        public void UpdatePictureBox(Image img)
        {
            logoPictureBox.Image = img;
        }

        public void UpdatePictureBox(string url)
        {
            logoPictureBox.ImageLocation = url;
        }
        public abstract void ShowAdvancedSettings();
        public abstract void OpenRepository();

        private void btnAdvancedSettings_Click(object sender, EventArgs e)
        {
            ShowAdvancedSettings();
        }

        /*
        //Método de Ejemplo Para Abrir Enlace de Repositorio
        public void OpenRepository()
        {
            try
            {
                Process.Start(Settings.CurrentSettings.DefaultRepositoryFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    @"No se puede abrir la ruta especificada. Asegúrese que la ruta es correcta e intente de nuevo." +
                    ex.Message, @"Abrir directorio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        */
    }
}