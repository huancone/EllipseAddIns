using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using EllipseCommonsClassLibrary.Connections;

// ReSharper disable ConvertPropertyToExpressionBody

namespace EllipseCommonsClassLibrary
{
    public partial class AboutBoxExcelAddIn : Form
    {
        private static Assembly _addInAssembly;
        private int _indexSettings;

        public AboutBoxExcelAddIn()
        {
            _addInAssembly = Assembly.GetCallingAssembly();
            
            InitializeComponent();
            this.Text = string.Format("About {0}", AssemblyTitle);
            this.labelProductName.Text = AssemblyProduct;
            this.labelVersion.Text = string.Format("Version {0}", AssemblyVersion);
            this.labelCopyright.Text = AssemblyCopyright;
            this.labelCompanyName.Text = AssemblyCompany;
            this.textBoxDescription.Text = AssemblyDescription;
            this.labelDeveloper1.Text = AssemblyDeveloper1;
            this.labelDeveloper2.Text = AssemblyDeveloper2;
        }

        public AboutBoxExcelAddIn(string developerName1, string developerName2)
        {
            _addInAssembly = Assembly.GetCallingAssembly();

            InitializeComponent();
            this.Text = string.Format("About {0}", AssemblyTitle);
            this.labelProductName.Text = AssemblyProduct;
            this.labelVersion.Text = string.Format("Version {0}", AssemblyVersion);
            this.labelCopyright.Text = AssemblyCopyright;
            this.labelCompanyName.Text = AssemblyCompany;
            this.textBoxDescription.Text = AssemblyDescription;
            this.labelDeveloper1.Text = developerName1;
            this.labelDeveloper2.Text = developerName2;
        }

        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = _addInAssembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(_addInAssembly.CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return _addInAssembly.GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = _addInAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
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
                object[] attributes = _addInAssembly.GetCustomAttributes(typeof(AssemblyProductAttribute), false);
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
                object[] attributes = _addInAssembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
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
                object[] attributes = _addInAssembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        public string AssemblyDeveloper1
        {
            get { return "hernandezrhectorj@gmail.com"; }
        }
        public string AssemblyDeveloper2
        {
            get { return "huancone@gmail.com"; }
        }
        #endregion

        private void tableLayoutPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void labelProductName_Click(object sender, EventArgs e)
        {
            _indexSettings++;
            if (_indexSettings > 3)
            {
                new SettingsBox(AssemblyProduct).ShowDialog();
            }
        }

        private void btnRepository_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(Configuration.DefaultRepositoryFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"No se puede abrir la ruta especificada. Asegúrese que la ruta es correcta e intente de nuevo." + ex.Message, @"Abrir directorio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}
