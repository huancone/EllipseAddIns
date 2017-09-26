using System;
using System.Reflection;
using System.Windows.Forms;
using EllipseCommonsClassLibrary.Connections;
namespace EllipseCommonsClassLibrary
{
    partial class SettingsBox : Form
    {
        public SettingsBox(string productName)
        {
            InitializeComponent();
            this.Text = @"Opciones de Configuración";
            this.labelProductName.Text = productName;
        }

        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        private void SettingsBox_Load(object sender, EventArgs e)
        {
            cbDebugErrors.Checked = Debugger.DebugErrors;
            cbDebugWarnings.Checked = Debugger.DebugWarnings;
            cbDebugQueries.Checked = Debugger.DebugQueries;
            tbLocalDataPath.Text = Configuration.LocalDataPath;
            tbServiceFileNetworkUrl.Text = Configuration.UrlServiceFileLocation;
        }

        private void cbDebugErrors_CheckedChanged(object sender, EventArgs e)
        {
            Debugger.DebugErrors = cbDebugErrors.Checked;
        }

        private void cbDebugWarnings_CheckedChanged(object sender, EventArgs e)
        {
            Debugger.DebugWarnings = cbDebugWarnings.Checked;
        }

        private void cbDebugQueries_CheckedChanged(object sender, EventArgs e)
        {
            Debugger.DebugQueries = cbDebugQueries.Checked;
        }

        private void tbLocalDataPath_TextChanged(object sender, EventArgs e)
        {
            Configuration.LocalDataPath = tbLocalDataPath.Text;
        }

        private void okButton_Click(object sender, EventArgs e)
        {

        }

        private void btnGenerateFromNetwork_Click(object sender, EventArgs e)
        {
            try
            {
                Configuration.GenerateEllipseConfigurationXmlFile(tbServiceFileNetworkUrl.Text);
                MessageBox.Show("Se ha generado el archivo local de configuración de Ellipse a partir del archivo de red " + Configuration.UrlServiceFileLocation + Configuration.ConfigXmlFileName, "Generate Local Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Generate Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenerateDefault_Click(object sender, EventArgs e)
        {
            try
            {
                Configuration.GenerateEllipseConfigurationXmlFile();
                MessageBox.Show("Se ha generado el archivo local de configuración de Ellipse de forma predeterminada", "Generate Local Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteLocalEllipseSettings_Click(object sender, EventArgs e)
        {
            try
            {
                Configuration.DeleteEllipseConfigurationXmlFile();
                MessageBox.Show("Se ha eliminado el archivo local de configuración de Ellipse", "Delete Local Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Delete Local Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cbForceRegionConfig_CheckedChanged(object sender, EventArgs e)
        {
            Debugger.DebugErrors = cbDebugErrors.Checked;
        }
    }
}
