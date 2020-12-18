using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using SharedClassLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Utilities;
using Debugger = SharedClassLibrary.Utilities.Debugger;

namespace SharedClassLibrary.Ellipse.Forms
{
    internal partial class SettingsBox : Form
    {
        public SettingsBox(string productName)
        {
            InitializeComponent();
            Text = @"Opciones de Configuración";
            labelProductName.Text = productName;
        }

        private void SettingsBox_Load(object sender, EventArgs e)
        {
            cbDebugErrors.Checked = Debugger.DebugErrors;
            cbDebugWarnings.Checked = Debugger.DebugWarnings;
            cbDebugQueries.Checked = Debugger.DebugQueries;
            tbLocalDataPath.Text = Settings.CurrentSettings.LocalDataPath;
            tbServiceFileNetworkUrl.Text = Settings.CurrentSettings.ServiceFilePath;
            tbTnsNameUrl.Text = Settings.CurrentSettings.TnsnamesFilePath;
            cbForceRegionConfig.Checked = Debugger.ForceRegionalization;
            cbForceServerList.Checked = Settings.CurrentSettings.IsServiceListForced;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            Settings.CurrentSettings.LocalDataPath = tbLocalDataPath.Text;
            Debugger.DebugErrors = cbDebugErrors.Checked;
            Debugger.DebugWarnings = cbDebugWarnings.Checked;
            Debugger.DebugQueries = cbDebugQueries.Checked;
            Settings.CurrentSettings.ServiceFilePath = tbServiceFileNetworkUrl.Text;
            Debugger.DebugErrors = cbDebugErrors.Checked;
            Debugger.ForceRegionalization = cbForceRegionConfig.Checked;
            RuntimeConfigSettings.UpdateTnsUrlValue(tbTnsNameUrl.Text);
            Debugger.LocalDataPath = tbLocalDataPath.Text;

            if (!Settings.CurrentSettings.IsServiceListForced.Equals(cbForceServerList.Checked))
                MessageBox.Show(
                    @"La configuración de Forzar Lista de Servidores ha cambiado. Debe reinicar la aplicación para que los cambios surjan efecto");
            Settings.CurrentSettings.IsServiceListForced = cbForceServerList.Checked;
        }

        private void btnGenerateFromUrl_Click(object sender, EventArgs e)
        {
           
        }
        private void btnRestoreServiceFile_Click(object sender, EventArgs e)
        {
            try
            {
                var dialogResult = MessageBox.Show("Se restaurará la ruta para los servicios a la ruta predeterminada. ¿Está seguro que desea continuar?", "Restore Ellipse Service File", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.No)
                    return;

                //Settings.CurrentSettings.DeleteEllipseConfigurationXmlFile();
                //restablecemos al valor predeterminado
                Settings.CurrentSettings.ServiceFilePath = Settings.CurrentSettings.DefaultServiceFilePath;
                tbServiceFileNetworkUrl.Text = Settings.CurrentSettings.ServiceFilePath;
                MessageBox.Show(
                    @"Se ha restaurado la dirección del archivo de configuración de los Servicios de Ellipse",
                    @"Restore Ellipse Service File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Restore Ellipse Service File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnGenerateServiceFileFromDefault_Click(object sender, EventArgs e)
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(tbServiceFileNetworkUrl.Text, true);
                var configFileName = Settings.CurrentSettings.ServicesConfigXmlFileName;
                var fileUrl = Path.Combine(configFilePath, configFileName);

                var dialogResult = MessageBox.Show("Se generará un nuevo archivo en la ruta especificada. ¿Está seguro que desea continuar?", "Generate Local Ellipse Settings File", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                    return;

                if (FileWriter.CheckFileExist(fileUrl))
                {
                    dialogResult = MessageBox.Show("El archivo local " + configFileName + " ya existe en la ruta especificada. ¿Está seguro que desea reemplazarlo?", "Generate Local Ellipse Settings File", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.No)
                        return;

                }

                Settings.CurrentSettings.GenerateConfigurationXmlFile(tbServiceFileNetworkUrl.Text);
                MessageBox.Show(@"Se ha generado el archivo local de configuración de Ellipse de forma predeterminada",
                    @"Generate Local Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Generate Ellipse Settings File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void btnGenerateServiceFileFromUrl_Click(object sender, EventArgs e)
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(tbServiceFileNetworkUrl.Text, true);
                var sourceFilePath = FileWriter.NormalizePath(Settings.CurrentSettings.DefaultServiceFilePath, true);
                var configFileName = Settings.CurrentSettings.ServicesConfigXmlFileName;
                var fileUrl = Path.Combine(configFilePath, configFileName);
                var sourceUrl = Path.Combine(sourceFilePath, configFileName);
                var dialogResult = MessageBox.Show("Se copiará el archivo de la ruta de red a la ruta especificada. ¿Está seguro que desea continuar?", "Generate Local Ellipse Settings File", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                    return;

                if (FileWriter.CheckFileExist(fileUrl))
                {
                    dialogResult = MessageBox.Show("El archivo local " + configFileName + " ya existe en la ruta especificada. ¿Está seguro que desea reemplazarlo?", "Generate Local Ellipse Settings File", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.No)
                        return;

                }

                Settings.CurrentSettings.GenerateConfigurationXmlFile(sourceFilePath, configFilePath);
                MessageBox.Show(@"Se ha generado el archivo local de configuración de Ellipse a partir del archivo de red " + sourceFilePath +
                    @" al directorio especificado", @"Generate Local Ellipse Settings File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Generate Ellipse Settings File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void btnRestoreLocalPath_Click(object sender, EventArgs e)
        {
            var dialogResult = MessageBox.Show("Se restaurará el directorio especificado al directorio predeterminado. ¿Está seguro que desea continuar?", "Restore Local Path", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No)
                return;

            tbLocalDataPath.Text = Settings.CurrentSettings.DefaultLocalDataPath;
        }

        private void btnOpenTnsnamesPath_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(tbTnsNameUrl.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    @"No se puede abrir la ruta especificada. Asegúrese que la ruta es correcta e intente de nuevo." +
                    ex.Message, @"Abrir directorio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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

        private void btnOpenServicesPath_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(tbServiceFileNetworkUrl.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    @"No se puede abrir la ruta especificada. Asegúrese que la ruta es correcta e intente de nuevo." +
                    ex.Message, @"Abrir directorio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        private void btnGenerateTnsnamesFile_Click(object sender, EventArgs e)
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(tbTnsNameUrl.Text, true);
                var configFileName = Settings.CurrentSettings.TnsnamesFileName;
                var fileUrl = Path.Combine(configFilePath, configFileName);
                if (FileWriter.CheckFileExist(fileUrl))
                {
                    var dialogResult = MessageBox.Show("El archivo local " + configFileName + " ya existe en la ruta especificada. ¿Está seguro que desea reemplazarlo?", "Generate Ellipse Tnsnames File", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.No)
                        return;

                }

                RuntimeConfigSettings.UpdateTnsUrlValue(tbTnsNameUrl.Text);
                Settings.CurrentSettings.GenerateEllipseTnsnamesFile(configFilePath);
                MessageBox.Show("Se ha generado el archivo local de TNSNAMES en la siguiente ruta: \n" + fileUrl,
                    @"Generate Local Ellipse Tnsnames File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Generate Ellipse Tnsnames File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnRestoreTnsnamesUrl_Click(object sender, EventArgs e)
        {
            try
            {
                var dialogResult = MessageBox.Show("Se restaurará la ruta para el TNSNAMES a la ruta predeterminada. ¿Está seguro que desea continuar?", "Restore Ellipse Tnsnames File", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.No)
                    return;

                RuntimeConfigSettings.UpdateTnsUrlValue(Settings.CurrentSettings.DefaultTnsnamesFilePath);
                tbTnsNameUrl.Text = Settings.CurrentSettings.DefaultTnsnamesFilePath;
                MessageBox.Show(
                    @"Se ha restaurado la ruta del archivo de TNSNAMES a su ubicación predeterminada. En caso de que este archivo no exista comuníquese con el administrador de sistemas",
                    @"Restore Ellipse Tnsnames File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, @"Restore Ellipse Tnsnames File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnGenerateCustomDb_Click(object sender, EventArgs e)
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(tbLocalDataPath.Text, true);
                var configFileName = Settings.CurrentSettings.DatabaseXmlFileName;
                var fileUrl = Path.Combine(configFilePath, configFileName);
                if (FileWriter.CheckFileExist(fileUrl))
                {
                    var dialogResult = MessageBox.Show("El archivo local de bases de datos de Ellipse ya existe en la ruta especificada. ¿Está seguro que desea reemplazarlo?", "Generate Local Ellipse Database File", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.No)
                        return;

                }
                Settings.CurrentSettings.GenerateDatabaseFile(configFilePath);

                MessageBox.Show(@"Se ha generado el archivo local de bases de datos de Ellipse en la ruta " + fileUrl,
                    @"Generate Local Ellipse Database File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se ha producido un error al intentar generar el archivo de base de datos. " + ex.Message, @"Generate Local Ellipse Database File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            
        }

        private void btnDeleteCustomDb_Click(object sender, EventArgs e)
        {
            try
            {
                var configFilePath = FileWriter.NormalizePath(tbLocalDataPath.Text, true);
                var configFileName = Settings.CurrentSettings.DatabaseXmlFileName;
                var fileUrl = Path.Combine(configFilePath, configFileName);
                if (FileWriter.CheckFileExist(fileUrl))
                {
                    var dialogResult = MessageBox.Show("Se eliminará el archivo local de bases de datos de Ellipse en la ruta especificada.\n" + fileUrl + "\n\n ¿Está seguro que desea eliminarlo?", "Generate Local Ellipse Database File", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.No)
                        return;

                }

                Settings.CurrentSettings.DeleteDatabaseFile(configFilePath);
                MessageBox.Show(@"Se ha eliminado el archivo local de bases de datos de Ellipse",
                    @"Delete Local Ellipse Database File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se ha producido un error al intentar eliminar el archivo de base de datos. " + ex.Message, @"Delete Local Ellipse Database File", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}