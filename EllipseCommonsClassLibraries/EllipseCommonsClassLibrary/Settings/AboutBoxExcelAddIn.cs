using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

// ReSharper disable ConvertPropertyToExpressionBody

namespace CommonsClassLibrary
{
    public abstract partial class AboutBoxExcelAddIn : Form
    {
        private int _indexSettings;
        private Settings.AssemblyItem addinAssembly;
        public AboutBoxExcelAddIn()
        {
            //addinAssembly = new Settings.AssemblyItem(Assembly.GetCallingAssembly());
            addinAssembly = new Settings.AssemblyItem(Settings.GetLastAssembly());
            InitializeComponent();
            Text = string.Format("About {0}", addinAssembly.AssemblyTitle);
            labelProductName.Text = addinAssembly.AssemblyProduct;
            labelVersion.Text = string.Format("Version {0}", addinAssembly.AssemblyVersion);
            labelCopyright.Text = addinAssembly.AssemblyCopyright;
            labelCompanyName.Text = addinAssembly.AssemblyCompany;
            textBoxDescription.Text = addinAssembly.AssemblyDescription;
            labelDeveloper1.Text = addinAssembly.AssemblyDeveloper1;
            labelDeveloper2.Text = addinAssembly.AssemblyDeveloper2;
        }

        public AboutBoxExcelAddIn(string developerName1, string developerName2)
        {
            //addinAssembly = new Settings.AssemblyItem(Assembly.GetCallingAssembly());
            addinAssembly = new Settings.AssemblyItem(Settings.GetLastAssembly());

            InitializeComponent();
            Text = string.Format("About {0}", addinAssembly.AssemblyTitle);
            labelProductName.Text = addinAssembly.AssemblyProduct;
            labelVersion.Text = string.Format("Version {0}", addinAssembly.AssemblyVersion);
            labelCopyright.Text = addinAssembly.AssemblyCopyright;
            labelCompanyName.Text = addinAssembly.AssemblyCompany;
            textBoxDescription.Text = addinAssembly.AssemblyDescription;
            labelDeveloper1.Text = developerName1;
            labelDeveloper2.Text = developerName2;
            //logoPictureBox.Image = Resources.ResourceManager.GetObject();
                
        }

        private void tableLayoutPanel_Paint(object sender, PaintEventArgs e)
        {
        }

        private void labelProductName_Click(object sender, EventArgs e)
        {
            _indexSettings++;
            if (_indexSettings > 3)
                ShowAdditionalOptions();
        }

        private void btnRepository_Click(object sender, EventArgs e)
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

        public void UpdatePictureBox(Image img)
        {
            logoPictureBox.Image = img;
        }

        public void UpdatePictureBox(string url)
        {
            logoPictureBox.ImageLocation = url;
        }
        public abstract void ShowAdditionalOptions();

    }
}