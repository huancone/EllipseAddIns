using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;

namespace LogsheetDatamodelAdmin
{
    public class AboutBox : SharedClassLibrary.Forms.AboutBoxExcelAddIn
    {
        public AboutBox() : base()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceFile = "EllipseCommonsClassLibrary.Resources.aboutPictureBox.png";
            var image = Image.FromStream(assembly.GetManifestResourceStream(resourceFile));

            UpdatePictureBox(image);
        }

        public AboutBox(string developer1, string developer2) : base(developer1, developer2)
        {
        }

        public override void ShowAdditionalOptions()
        {
            throw new NotImplementedException();
        }

        public override void OpenRepository()
        {
            throw new NotImplementedException();
        }
    }
}
