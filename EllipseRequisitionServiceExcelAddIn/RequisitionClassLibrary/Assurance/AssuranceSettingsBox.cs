using System;
using System.Diagnostics;
using System.Windows.Forms;
using Settings = EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.Assurance.AssuranceSettings;

namespace EllipseRequisitionServiceExcelAddIn.RequisitionClassLibrary.Assurance
{
    internal partial class AssuranceSettingsBox : Form
    {
        Settings.Configuration config;

        public AssuranceSettingsBox()
        {
            try
            {
                config = Settings.GetSettings();
            }
            catch(System.IO.FileNotFoundException ex)
            {
                config = Settings.CreateSettingsFile();
                SharedClassLibrary.Utilities.Debugger.LogError("AssuranceSettingsBox.cs:AssuranceSettingsBox()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
            }
            catch(System.IO.DirectoryNotFoundException ex)
            {
                config = Settings.CreateSettingsFile();
                SharedClassLibrary.Utilities.Debugger.LogError("AssuranceSettingsBox.cs:AssuranceSettingsBox()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
            }
            catch(Exception ex)
            {
                SharedClassLibrary.Utilities.Debugger.LogError("AssuranceSettingsBox.cs:AssuranceSettingsBox()", "\n\rMessage: " + ex.Message + "\n\rSource: " + ex.Source + "\n\rStackTrace: " + ex.StackTrace);
                MessageBox.Show(@"Se ha producido un error: " + ex.Message);
            }
            InitializeComponent();
        }

        private void SettingsBox_Load(object sender, EventArgs e)
        {
            tbLocalDataPath.Text = Settings.LocalDataPath;
            tbMinItemValue.Text = "" + config.minItemValue;
            cbCostCenterMatch.Checked = config.costCenterMatch;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            Settings.LocalDataPath = tbLocalDataPath.Text;
            config.minItemValue = Convert.ToInt32(tbMinItemValue.Text);
            config.costCenterMatch = cbCostCenterMatch.Checked;
            Settings.UpdateSettings(config);
        }

        private void btnRestoreLocalPath_Click(object sender, EventArgs e)
        {
            tbLocalDataPath.Text = Settings.DefaultLocalDataPath;
        }

        private void btnOpenLocalPath_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(tbLocalDataPath.Text);
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