using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EllipseSubAssetGeneralInfoExcelAddIn
{
    public partial class FormAuthenticate : Form
    {
        public FormAuthenticate()
        {
            InitializeComponent();
            txtDistrict.Text = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.DstrDefault;
            txtPosition.Text = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.PosDefault;
            txtUsername.Text = global::EllipseSubAssetGeneralInfoExcelAddIn.Properties.Settings.Default.UserDefault;
        }

        private void btnAuthenticate_Click(object sender, EventArgs e)
        {
            RibbonEllipse.ElliseDsct = txtDistrict.Text.ToUpper();
            RibbonEllipse.EllisePost = txtPosition.Text.ToUpper();
            RibbonEllipse.EllisePswd = txtPassword.Text;
            RibbonEllipse.ElliseUser = txtUsername.Text.ToUpper();
            clearForm();
            this.Close();
        }

        public void clearForm()
        {
            txtDistrict.Clear();
            txtPassword.Clear();
            txtPosition.Clear();
            txtUsername.Clear();
        }
    }
}
