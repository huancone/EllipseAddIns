using System;
using System.Reflection;
using System.Windows.Forms;

namespace EllipseCommonsClassLibrary
{
    partial class SettingsBox : Form
    {
        public SettingsBox()
        {
            InitializeComponent();
            this.Text = @"Opciones de Configuración";
            this.labelProductName.Text = AssemblyProduct;
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
            tbLocalDataPath.Text = Debugger.LocalDataPath;
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
            Debugger.LocalDataPath = tbLocalDataPath.Text;
        }
    }
}
