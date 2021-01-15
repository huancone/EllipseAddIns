using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using erl.Oracle.TnsNames;
using SharedClassLibrary;
using SharedClassLibrary.Configuration;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Utilities.Encryption;
using SharedClassLibrary.Ellipse.Forms;
using SharedClassLibrary.Utilities;

namespace SharedClassUtility
{
    public partial class MainForm : Form
    {
        private TnsNameInfo[] _tnsNames;
        private EllipseFunctions _eFunctions;
        private FormAuthenticate _frmAuth;

        public MainForm()
        {
            InitializeComponent();
            tbDbName.Text = "EL8PROD";
            tbDbUser.Text = "SIGCON";
            tbDbPassword.Text = "ventyx";

            cbOraSource.Items.Add(TnsNames.GetSourceName(TnsNamesSource.OracleHomeEnvironmentVariable));
            cbOraSource.Items.Add(TnsNames.GetSourceName(TnsNamesSource.TnsAdminEnvironmentVariable));
            cbOraSource.Items.Add(TnsNames.GetSourceName(TnsNamesSource.PathEnvironmentVariable));
            cbOraSource.Items.Add(TnsNames.GetSourceName(TnsNamesSource.CustomPath));
        }

        private void btnEncrypt_Click(object sender, EventArgs e)
        {
            try
            {
                EncryptText();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERROR");
            }
        }

        private void btnDecrypt_Click(object sender, EventArgs e)
        {
            try
            {
                DecryptText();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERROR");
            }
        }

        public void EncryptText()
        {
            var text = tbPlainText.Text.Trim();
            var passPhrase = string.IsNullOrWhiteSpace(tbPassPhrase.Text) ? null : tbPassPhrase.Text.Trim();

            var result = Encryption.Encrypt(text, passPhrase);
            tbCipherText.Text = result;
        }

        public void DecryptText()
        {
            var text = tbPlainText.Text.Trim();
            var passPhrase = string.IsNullOrWhiteSpace(tbPassPhrase.Text) ? null : tbPassPhrase.Text.Trim();

            var result = Encryption.Decrypt(text, passPhrase);
            tbCipherText.Text = result;
        }
        private void btnCopyCipherText_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(tbCipherText.Text);
        }

        private void btnPastePlainText_Click(object sender, EventArgs e)
        {
            tbPlainText.Text = Clipboard.GetText();
        }

        private void btnPastePassPhrase_Click(object sender, EventArgs e)
        {
            tbPassPhrase.Text = Clipboard.GetText();
        }

        private void btnCleanPlainText_Click(object sender, EventArgs e)
        {
            tbPlainText.Text = "";
        }

        private void btnCleanPassPhrase_Click(object sender, EventArgs e)
        {
            tbPassPhrase.Text = "";
        }

        public void TestDatabaseConnection()
        {
            try
            {
                var connectionString = "";
                if (tabcDbConnectionMode.SelectedTab == tabConnectionString)
                    connectionString = tbConnectionString.Text;

                var dbName = tbDbName.Text;
                var dbUser = tbDbUser.Text;
                var dbPassword = tbDbPassword.Text;
                var dbCipherPassword = tbDbCipheredPassword.Text;
                var dbType = cbDatabaseType.Text;

                var dbItem = new DatabaseItem();

                dbItem.DbName = dbName;
                dbItem.DbUser = dbUser;
                if (!string.IsNullOrWhiteSpace(dbPassword))
                    dbItem.DbPassword = dbPassword;
                if (!string.IsNullOrWhiteSpace(dbCipherPassword))
                    dbItem.DbEncodedPassword = dbCipherPassword;

                string testQuery;
                IDbConnector connector;


                if (dbType.Equals("ORACLE", StringComparison.InvariantCultureIgnoreCase))
                {
                    testQuery = "SELECT 1 AS NUMBERVALUE FROM DUAL";

                    if (_tnsNames != null)
                    {
                        foreach (var tns in _tnsNames)
                        {
                            if (!dbItem.DbName.Equals(tns.TnsName)) continue;
                            dbItem.DbName = tns.DataSource;
                            break;
                        }
                    }

                    connector = string.IsNullOrWhiteSpace(connectionString) ? new OracleConnector(dbItem) : new OracleConnector(connectionString);
                }
                else if (dbType.Equals("SQLSERVER", StringComparison.InvariantCultureIgnoreCase))
                {
                    testQuery = "SELECT 1 AS NUMBERVALUE";
                    connector = string.IsNullOrWhiteSpace(connectionString) ? new SqlConnector(dbItem) : new SqlConnector(connectionString);
                }
                else
                {
                    throw new Exception("INVALID DATABASE TYPE");
                }



                connector.StartConnection();
                var result = connector.ExecuteQuery(testQuery);
                MessageBox.Show("Connected", "Database Connection Test");
                connector.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection Failed: " + ex.Message);
            }
        }

        private void btnTestDbConnection_Click_1(object sender, EventArgs e)
        {
            TestDatabaseConnection();
        }

        private void btnExecuteQuery_Click(object sender, EventArgs e)
        {
            ExecuteQuery();
        }

        private void ExecuteQuery()
        {
            try
            {
                var connectionString = "";
                if (tabcDbConnectionMode.SelectedTab == tabConnectionString)
                    connectionString = tbConnectionString.Text;

                var dbName = tbDbName.Text;
                var dbUser = tbDbUser.Text;
                var dbPassword = tbDbPassword.Text;
                var dbCipherPassword = tbDbCipheredPassword.Text;
                var dbType = cbDatabaseType.Text;

                var dbItem = new DatabaseItem();

                dbItem.DbName = dbName;
                dbItem.DbUser = dbUser;
                if (!string.IsNullOrWhiteSpace(dbPassword))
                    dbItem.DbPassword = dbPassword;
                if (!string.IsNullOrWhiteSpace(dbCipherPassword))
                    dbItem.DbEncodedPassword = dbCipherPassword;

                var testQuery = tbDbQuery.Text;
                IDbConnector connector;


                if (dbType.Equals("ORACLE", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (_tnsNames != null)
                    {
                        foreach (var tns in _tnsNames)
                        {
                            if (!dbItem.DbName.Equals(tns.TnsName)) continue;
                            dbItem.DbName = tns.DataSource;
                            break;
                        }
                    }

                    connector = string.IsNullOrWhiteSpace(connectionString) ? new OracleConnector(dbItem) : new OracleConnector(connectionString);
                }
                else if (dbType.Equals("SQLSERVER", StringComparison.InvariantCultureIgnoreCase))
                {
                    connector = string.IsNullOrWhiteSpace(connectionString) ? new SqlConnector(dbItem) : new SqlConnector(connectionString);
                }
                else
                {
                    throw new Exception("INVALID DATABASE TYPE");
                }



                connector.StartConnection();
                var result = connector.GetQueryResult(testQuery);

                var outputText = "";

                if (result == null)
                    return;
                
                //Cargo el encabezado de la tabla y doy formato
                for (var i = 0; i < result.FieldCount; i++)
                    outputText+= "\n" + result.GetName(i) + ";";

                //cargo los datos 
                if (result.IsClosed) return;

                while (result.Read())
                {
                    for (var i = 0; i < result.FieldCount; i++)
                        outputText += "\n" + result[i].ToString().Trim() + ";";
                }

                MessageBox.Show(outputText, "Executed Query");
                connector.CloseConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection Error: " + ex.Message);
            }
        }

        private void cbOraSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            OracleSourceItemChanged();
        }

        private void OracleSourceItemChanged()
        {
            cbOraServers.Items.Clear();
            _tnsNames = null;
            try
            {
                if (cbOraSource.SelectedItem.Equals(TnsNames.GetSourceName(TnsNamesSource.CustomPath)))
                {
                    if (!string.IsNullOrWhiteSpace(tbOraFilePath.Text))
                        _tnsNames = TnsNames.OpenTnsNameInfo(tbOraFilePath.Text);
                }
                else if (cbOraSource.SelectedItem.Equals(TnsNames.GetSourceName(TnsNamesSource.OracleHomeEnvironmentVariable)))
                    _tnsNames = TnsNames.OpenTnsNameInfo(TnsNamesSource.OracleHomeEnvironmentVariable);
                else if (cbOraSource.SelectedItem.Equals(TnsNames.GetSourceName(TnsNamesSource.TnsAdminEnvironmentVariable)))
                    _tnsNames = TnsNames.OpenTnsNameInfo(TnsNamesSource.TnsAdminEnvironmentVariable);
                else if (cbOraSource.SelectedItem.Equals(TnsNames.GetSourceName(TnsNamesSource.PathEnvironmentVariable)))
                    _tnsNames = TnsNames.OpenTnsNameInfo(TnsNamesSource.PathEnvironmentVariable);

                if (_tnsNames == null) return;
                foreach (var tns in _tnsNames)
                    cbOraServers.Items.Add(tns.TnsName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SOURCE ERROR");
            }
        }
        private void cbOraServers_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbDbName.Text = cbOraServers.SelectedItem.ToString();
        }

        private void OracleFilePathChanged()
        {
            try
            {
                if (cbOraSource.SelectedItem == null)
                    return;
                if (!cbOraSource.SelectedItem.Equals(TnsNames.GetSourceName(TnsNamesSource.CustomPath))) return;
                cbOraServers.Items.Clear();
                
                if (!string.IsNullOrWhiteSpace(tbOraFilePath.Text))
                    _tnsNames = TnsNames.OpenTnsNameInfo(tbOraFilePath.Text);
                if (_tnsNames == null) return;
                foreach (var tns in _tnsNames)
                    cbOraServers.Items.Add(tns.TnsName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ORA FILE PATH ERROR");
            }
        }
        private void cbDatabaseType_SelectedIndexChanged(object sender, EventArgs e)
        {
            OracleItemsStatusChanged();
        }

        private void OracleItemsStatusChanged()
        {
            if (cbDatabaseType.SelectedItem.Equals("ORACLE"))
            {
                cbOraSource.Enabled = true;
                cbOraServers.Enabled = true;
                tbOraFilePath.Enabled = true;

                cbOraSource.Visible = true;
                cbOraServers.Visible = true;
                tbOraFilePath.Visible = true;

                lblOraSource.Visible = true;
                lblOraServers.Visible = true;
                lblOraFilePath.Visible = true;
            }
            else
            {

                cbOraSource.Enabled = false;
                cbOraServers.Enabled = false;
                tbOraFilePath.Enabled = false;

                cbOraSource.Visible = false;
                cbOraServers.Visible = false;
                tbOraFilePath.Visible = false;

                lblOraSource.Visible = false;
                lblOraServers.Visible = false;
                lblOraFilePath.Visible = false;
            }
        }

        private void tbOraFilePath_Leave(object sender, EventArgs e)
        {
            OracleFilePathChanged();
        }

        private void btnEllipseAbout_Click(object sender, EventArgs e)
        {
            new SharedClassLibrary.Ellipse.Forms.AboutBoxExcelAddIn().ShowDialog();
        }

        private void btnEllipseSettings_Click(object sender, EventArgs e)
        {
            new SharedClassLibrary.Ellipse.Forms.AboutBoxExcelAddIn().ShowAdditionalOptions();
        }

        private void btnStartEllipseSettings_Click(object sender, EventArgs e)
        {
            LoadEllipseSettings();
        }

        public void LoadEllipseSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();
            _frmAuth = new FormAuthenticate();

            var environments = Environments.GetEnvironmentList();
            drpEnvironment.Items.Clear();
            foreach (var env in environments)
                drpEnvironment.Items.Add(env);

            //settings.SetDefaultCustomSettingValue("OptionName1", "false");
            //settings.SetDefaultCustomSettingValue("OptionName2", "OptionValue2");
            //settings.SetDefaultCustomSettingValue("OptionName3", "OptionValue3");



            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //var optionItem1Value = MyUtilities.IsTrue(settings.GetCustomSettingValue("OptionName1"));
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName2");
            //var optionItem1Value = settings.GetCustomSettingValue("OptionName3");

            //cbCustomSettingOption.Checked = optionItem1Value;
            //optionItem2.Text = optionItem2Value;
            //optionItem3 = optionItem3Value;

            //
            settings.SaveCustomSettings();
        }
    }
}
