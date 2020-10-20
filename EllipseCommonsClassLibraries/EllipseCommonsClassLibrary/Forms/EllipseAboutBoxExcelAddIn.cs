using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Reflection;
using EllipseCommonsClassLibrary.Properties;
using System.Windows.Forms;
using System.Diagnostics;

namespace EllipseCommonsClassLibrary
{
    public class AboutBoxExcelAddIn : SharedClassLibrary.Forms.AboutBoxExcelAddIn
    {
        public AboutBoxExcelAddIn() : base()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceFile = "EllipseCommonsClassLibrary.Resources.aboutPictureBox.png";
            var image = Image.FromStream(assembly.GetManifestResourceStream(resourceFile));

            UpdatePictureBox(image);
        }

        public AboutBoxExcelAddIn(string developer1, string developer2) : base(developer1, developer2)
        {
        }

        override 
        public void ShowAdditionalOptions()
        {
            var commonAssembly = new Settings.AssemblyItem(Assembly.GetExecutingAssembly());
            var productLabel = commonAssembly.AssemblyProduct + " v" + commonAssembly.AssemblyVersion;
            new SettingsBox(productLabel).ShowDialog();
        }

        public override void OpenRepository()
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
    }
}
