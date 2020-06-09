using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using EllipseCommonsClassLibrary.Connections;
using EllipseCommonsClassLibrary.Settings;

// ReSharper disable ConvertPropertyToExpressionBody

namespace EllipseCommonsClassLibrary
{
    public partial class AboutBoxExcelAddIn : Form
    {
        private int _indexSettings;
        private AssemblyItem addinAssembly;
        public AboutBoxExcelAddIn()
        {
            addinAssembly = new AssemblyItem(Assembly.GetCallingAssembly());

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
            addinAssembly = new AssemblyItem(Assembly.GetCallingAssembly());

            InitializeComponent();
            Text = string.Format("About {0}", addinAssembly.AssemblyTitle);
            labelProductName.Text = addinAssembly.AssemblyProduct;
            labelVersion.Text = string.Format("Version {0}", addinAssembly.AssemblyVersion);
            labelCopyright.Text = addinAssembly.AssemblyCopyright;
            labelCompanyName.Text = addinAssembly.AssemblyCompany;
            textBoxDescription.Text = addinAssembly.AssemblyDescription;
            labelDeveloper1.Text = developerName1;
            labelDeveloper2.Text = developerName2;
        }

        private void tableLayoutPanel_Paint(object sender, PaintEventArgs e)
        {
        }

        private void labelProductName_Click(object sender, EventArgs e)
        {
            //var productLabel = addinAssembly.AssemblyProduct;
            var commonAssembly = new AssemblyItem(Assembly.GetExecutingAssembly());
            var productLabel = commonAssembly.AssemblyProduct + " v" + commonAssembly.AssemblyVersion;
            _indexSettings++;
            if (_indexSettings > 3) new SettingsBox(productLabel).ShowDialog();
        }

        private void btnRepository_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(Configuration.DefaultRepositoryFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    @"No se puede abrir la ruta especificada. Asegúrese que la ruta es correcta e intente de nuevo." +
                    ex.Message, @"Abrir directorio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}