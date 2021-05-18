using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using SharedClassLibrary.Ellipse.Properties;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using SharedClassLibrary.Configuration;
using SharedClassLibrary.Ellipse.Forms;

namespace SharedClassLibrary.Ellipse.Forms
{
    public class AboutBoxExcelAddIn : SharedClassLibrary.Forms.AboutBoxExcelAddIn
    {
        public AboutBoxExcelAddIn() : base()
        {
            var assembly = Assembly.GetExecutingAssembly();
            
            //var resourceFile = assembly.GetName().Name + ".Resources.aboutPictureBox.png";
            var image = Ellipse.EllipseResources.aboutPictureBox;//Image.FromStream(assembly.GetManifestResourceStream(resourceFile));

            UpdatePictureBox(image);
            ShowAdvancedSettingsButton(true);
        }
        public AboutBoxExcelAddIn(string developer1) : base(developer1, null)
        {
            ShowAdvancedSettingsButton(true);
        }
        public AboutBoxExcelAddIn(string developer1, bool showAdvancedSettingsButton) : base(developer1, null)
        {
            ShowAdvancedSettingsButton(showAdvancedSettingsButton);
        }
        public AboutBoxExcelAddIn(string developer1, string developer2) : base(developer1, developer2)
        {
            ShowAdvancedSettingsButton(true);
        }
        public AboutBoxExcelAddIn(string developer1, string developer2, bool showAdvancedSettingsButton) : base(developer1, developer2)
        {
            ShowAdvancedSettingsButton(showAdvancedSettingsButton);
        }
        override 
        public void ShowAdvancedSettings()
        {
            var commonAssembly = new AssemblyItem(Assembly.GetExecutingAssembly());
            var productLabel = commonAssembly.AssemblyProduct + " v" + commonAssembly.AssemblyVersion;
            new SettingsBox(productLabel).ShowDialog();
        }

        public override void OpenRepository()
        {
            try
            {
                Process.Start(SharedClassLibrary.Ellipse.Settings.CurrentSettings.DefaultRepositoryFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $@"No se puede abrir la ruta especificada. Asegúrese que la ruta es correcta e intente de nuevo. 
                          {SharedClassLibrary.Ellipse.Settings.CurrentSettings.DefaultRepositoryFilePath}. {ex.Message}", 
                    @"Abrir directorio", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
