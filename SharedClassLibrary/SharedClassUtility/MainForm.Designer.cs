
namespace SharedClassUtility
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabcGeneral = new System.Windows.Forms.TabControl();
            this.tabHome = new System.Windows.Forms.TabPage();
            this.lblDevelopBy = new System.Windows.Forms.Label();
            this.lblDeveloper = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.tabEncryption = new System.Windows.Forms.TabPage();
            this.btnCleanPlainText = new System.Windows.Forms.Button();
            this.btnCleanPassPhrase = new System.Windows.Forms.Button();
            this.btnCopyCipherText = new System.Windows.Forms.Button();
            this.btnPastePassPhrase = new System.Windows.Forms.Button();
            this.btnPastePlainText = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.btnDecrypt = new System.Windows.Forms.Button();
            this.tbCipherText = new System.Windows.Forms.TextBox();
            this.lblPassPhrase = new System.Windows.Forms.Label();
            this.lblEncryptText = new System.Windows.Forms.Label();
            this.tbPassPhrase = new System.Windows.Forms.TextBox();
            this.tbPlainText = new System.Windows.Forms.TextBox();
            this.btnEncrypt = new System.Windows.Forms.Button();
            this.tabDbConnection = new System.Windows.Forms.TabPage();
            this.lblDbType = new System.Windows.Forms.Label();
            this.cbDatabaseType = new System.Windows.Forms.ComboBox();
            this.tabcDbConnectionMode = new System.Windows.Forms.TabControl();
            this.tabConnectionString = new System.Windows.Forms.TabPage();
            this.tbConnectionString = new System.Windows.Forms.TextBox();
            this.lblConnectionString = new System.Windows.Forms.Label();
            this.tabConnectionItem = new System.Windows.Forms.TabPage();
            this.gbDatabaseConnection = new System.Windows.Forms.GroupBox();
            this.lblOraSource = new System.Windows.Forms.Label();
            this.cbOraSource = new System.Windows.Forms.ComboBox();
            this.lblOraServers = new System.Windows.Forms.Label();
            this.cbOraServers = new System.Windows.Forms.ComboBox();
            this.tbOraFilePath = new System.Windows.Forms.TextBox();
            this.lblOraFilePath = new System.Windows.Forms.Label();
            this.lblCipheredPassword = new System.Windows.Forms.Label();
            this.lblDbPassword = new System.Windows.Forms.Label();
            this.lblDbUser = new System.Windows.Forms.Label();
            this.tbDbCipheredPassword = new System.Windows.Forms.TextBox();
            this.tbDbPassword = new System.Windows.Forms.TextBox();
            this.tbDbUser = new System.Windows.Forms.TextBox();
            this.tbDbName = new System.Windows.Forms.TextBox();
            this.lblDbName = new System.Windows.Forms.Label();
            this.gbQuery = new System.Windows.Forms.GroupBox();
            this.btnExecuteQuery = new System.Windows.Forms.Button();
            this.tbDbQuery = new System.Windows.Forms.TextBox();
            this.btnTestDbConnection = new System.Windows.Forms.Button();
            this.tabEllipseSettings = new System.Windows.Forms.TabPage();
            this.drpEnvironment = new System.Windows.Forms.ComboBox();
            this.lblEnvironment = new System.Windows.Forms.Label();
            this.btnStartEllipseSettings = new System.Windows.Forms.Button();
            this.btnEllipseSettings = new System.Windows.Forms.Button();
            this.btnEllipseAbout = new System.Windows.Forms.Button();
            this.tabcGeneral.SuspendLayout();
            this.tabHome.SuspendLayout();
            this.tabEncryption.SuspendLayout();
            this.tabDbConnection.SuspendLayout();
            this.tabcDbConnectionMode.SuspendLayout();
            this.tabConnectionString.SuspendLayout();
            this.tabConnectionItem.SuspendLayout();
            this.gbDatabaseConnection.SuspendLayout();
            this.gbQuery.SuspendLayout();
            this.tabEllipseSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabcGeneral
            // 
            this.tabcGeneral.Controls.Add(this.tabHome);
            this.tabcGeneral.Controls.Add(this.tabEncryption);
            this.tabcGeneral.Controls.Add(this.tabDbConnection);
            this.tabcGeneral.Controls.Add(this.tabEllipseSettings);
            this.tabcGeneral.Location = new System.Drawing.Point(12, 12);
            this.tabcGeneral.Name = "tabcGeneral";
            this.tabcGeneral.SelectedIndex = 0;
            this.tabcGeneral.Size = new System.Drawing.Size(440, 426);
            this.tabcGeneral.TabIndex = 0;
            // 
            // tabHome
            // 
            this.tabHome.Controls.Add(this.lblDevelopBy);
            this.tabHome.Controls.Add(this.lblDeveloper);
            this.tabHome.Controls.Add(this.lblTitle);
            this.tabHome.Location = new System.Drawing.Point(4, 22);
            this.tabHome.Name = "tabHome";
            this.tabHome.Padding = new System.Windows.Forms.Padding(3);
            this.tabHome.Size = new System.Drawing.Size(432, 400);
            this.tabHome.TabIndex = 0;
            this.tabHome.Text = "Home";
            this.tabHome.UseVisualStyleBackColor = true;
            // 
            // lblDevelopBy
            // 
            this.lblDevelopBy.AutoSize = true;
            this.lblDevelopBy.Location = new System.Drawing.Point(6, 127);
            this.lblDevelopBy.Name = "lblDevelopBy";
            this.lblDevelopBy.Size = new System.Drawing.Size(64, 13);
            this.lblDevelopBy.TabIndex = 2;
            this.lblDevelopBy.Text = "Develop by:";
            // 
            // lblDeveloper
            // 
            this.lblDeveloper.Location = new System.Drawing.Point(6, 149);
            this.lblDeveloper.Name = "lblDeveloper";
            this.lblDeveloper.Size = new System.Drawing.Size(180, 36);
            this.lblDeveloper.TabIndex = 1;
            this.lblDeveloper.Text = "Héctor J Hernández R <hernandezrhectorj@gmail.com>";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(6, 3);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(103, 13);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Shared Class Library";
            // 
            // tabEncryption
            // 
            this.tabEncryption.Controls.Add(this.btnCleanPlainText);
            this.tabEncryption.Controls.Add(this.btnCleanPassPhrase);
            this.tabEncryption.Controls.Add(this.btnCopyCipherText);
            this.tabEncryption.Controls.Add(this.btnPastePassPhrase);
            this.tabEncryption.Controls.Add(this.btnPastePlainText);
            this.tabEncryption.Controls.Add(this.lblResult);
            this.tabEncryption.Controls.Add(this.btnDecrypt);
            this.tabEncryption.Controls.Add(this.tbCipherText);
            this.tabEncryption.Controls.Add(this.lblPassPhrase);
            this.tabEncryption.Controls.Add(this.lblEncryptText);
            this.tabEncryption.Controls.Add(this.tbPassPhrase);
            this.tabEncryption.Controls.Add(this.tbPlainText);
            this.tabEncryption.Controls.Add(this.btnEncrypt);
            this.tabEncryption.Location = new System.Drawing.Point(4, 22);
            this.tabEncryption.Name = "tabEncryption";
            this.tabEncryption.Padding = new System.Windows.Forms.Padding(3);
            this.tabEncryption.Size = new System.Drawing.Size(432, 400);
            this.tabEncryption.TabIndex = 1;
            this.tabEncryption.Text = "Encryption";
            this.tabEncryption.UseVisualStyleBackColor = true;
            // 
            // btnCleanPlainText
            // 
            this.btnCleanPlainText.Image = global::SharedClassUtility.Properties.Resources.CleanData_disabled_16x;
            this.btnCleanPlainText.Location = new System.Drawing.Point(248, 21);
            this.btnCleanPlainText.Name = "btnCleanPlainText";
            this.btnCleanPlainText.Size = new System.Drawing.Size(37, 31);
            this.btnCleanPlainText.TabIndex = 13;
            this.btnCleanPlainText.UseVisualStyleBackColor = true;
            this.btnCleanPlainText.Click += new System.EventHandler(this.btnCleanPlainText_Click);
            // 
            // btnCleanPassPhrase
            // 
            this.btnCleanPassPhrase.Image = global::SharedClassUtility.Properties.Resources.CleanData_disabled_16x;
            this.btnCleanPassPhrase.Location = new System.Drawing.Point(248, 64);
            this.btnCleanPassPhrase.Name = "btnCleanPassPhrase";
            this.btnCleanPassPhrase.Size = new System.Drawing.Size(37, 31);
            this.btnCleanPassPhrase.TabIndex = 12;
            this.btnCleanPassPhrase.UseVisualStyleBackColor = true;
            this.btnCleanPassPhrase.Click += new System.EventHandler(this.btnCleanPassPhrase_Click);
            // 
            // btnCopyCipherText
            // 
            this.btnCopyCipherText.Image = global::SharedClassUtility.Properties.Resources.ASX_Copy_grey_16x;
            this.btnCopyCipherText.Location = new System.Drawing.Point(205, 139);
            this.btnCopyCipherText.Name = "btnCopyCipherText";
            this.btnCopyCipherText.Size = new System.Drawing.Size(37, 31);
            this.btnCopyCipherText.TabIndex = 11;
            this.btnCopyCipherText.UseVisualStyleBackColor = true;
            this.btnCopyCipherText.Click += new System.EventHandler(this.btnCopyCipherText_Click);
            // 
            // btnPastePassPhrase
            // 
            this.btnPastePassPhrase.Image = global::SharedClassUtility.Properties.Resources.ASX_Paste_grey_16x;
            this.btnPastePassPhrase.Location = new System.Drawing.Point(205, 64);
            this.btnPastePassPhrase.Name = "btnPastePassPhrase";
            this.btnPastePassPhrase.Size = new System.Drawing.Size(37, 31);
            this.btnPastePassPhrase.TabIndex = 10;
            this.btnPastePassPhrase.UseVisualStyleBackColor = true;
            this.btnPastePassPhrase.Click += new System.EventHandler(this.btnPastePassPhrase_Click);
            // 
            // btnPastePlainText
            // 
            this.btnPastePlainText.Image = global::SharedClassUtility.Properties.Resources.ASX_Paste_grey_16x;
            this.btnPastePlainText.Location = new System.Drawing.Point(205, 21);
            this.btnPastePlainText.Name = "btnPastePlainText";
            this.btnPastePlainText.Size = new System.Drawing.Size(37, 31);
            this.btnPastePlainText.TabIndex = 9;
            this.btnPastePlainText.UseVisualStyleBackColor = true;
            this.btnPastePlainText.Click += new System.EventHandler(this.btnPastePlainText_Click);
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(10, 126);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(37, 13);
            this.lblResult.TabIndex = 7;
            this.lblResult.Text = "Result";
            // 
            // btnDecrypt
            // 
            this.btnDecrypt.Location = new System.Drawing.Point(118, 97);
            this.btnDecrypt.Name = "btnDecrypt";
            this.btnDecrypt.Size = new System.Drawing.Size(81, 23);
            this.btnDecrypt.TabIndex = 6;
            this.btnDecrypt.Text = "&Decrypt";
            this.btnDecrypt.UseVisualStyleBackColor = true;
            this.btnDecrypt.Click += new System.EventHandler(this.btnDecrypt_Click);
            // 
            // tbCipherText
            // 
            this.tbCipherText.Location = new System.Drawing.Point(10, 145);
            this.tbCipherText.Name = "tbCipherText";
            this.tbCipherText.Size = new System.Drawing.Size(189, 20);
            this.tbCipherText.TabIndex = 5;
            // 
            // lblPassPhrase
            // 
            this.lblPassPhrase.AutoSize = true;
            this.lblPassPhrase.Location = new System.Drawing.Point(10, 54);
            this.lblPassPhrase.Name = "lblPassPhrase";
            this.lblPassPhrase.Size = new System.Drawing.Size(66, 13);
            this.lblPassPhrase.TabIndex = 4;
            this.lblPassPhrase.Text = "Pass Phrase";
            // 
            // lblEncryptText
            // 
            this.lblEncryptText.AutoSize = true;
            this.lblEncryptText.Location = new System.Drawing.Point(7, 11);
            this.lblEncryptText.Name = "lblEncryptText";
            this.lblEncryptText.Size = new System.Drawing.Size(125, 13);
            this.lblEncryptText.TabIndex = 3;
            this.lblEncryptText.Text = "Text To Encrypt/Decrypt";
            // 
            // tbPassPhrase
            // 
            this.tbPassPhrase.Location = new System.Drawing.Point(10, 70);
            this.tbPassPhrase.Name = "tbPassPhrase";
            this.tbPassPhrase.Size = new System.Drawing.Size(189, 20);
            this.tbPassPhrase.TabIndex = 2;
            // 
            // tbPlainText
            // 
            this.tbPlainText.Location = new System.Drawing.Point(10, 27);
            this.tbPlainText.Name = "tbPlainText";
            this.tbPlainText.Size = new System.Drawing.Size(189, 20);
            this.tbPlainText.TabIndex = 1;
            // 
            // btnEncrypt
            // 
            this.btnEncrypt.Location = new System.Drawing.Point(10, 96);
            this.btnEncrypt.Name = "btnEncrypt";
            this.btnEncrypt.Size = new System.Drawing.Size(86, 23);
            this.btnEncrypt.TabIndex = 0;
            this.btnEncrypt.Text = "&Encrypt";
            this.btnEncrypt.UseVisualStyleBackColor = true;
            this.btnEncrypt.Click += new System.EventHandler(this.btnEncrypt_Click);
            // 
            // tabDbConnection
            // 
            this.tabDbConnection.Controls.Add(this.lblDbType);
            this.tabDbConnection.Controls.Add(this.cbDatabaseType);
            this.tabDbConnection.Controls.Add(this.tabcDbConnectionMode);
            this.tabDbConnection.Controls.Add(this.gbQuery);
            this.tabDbConnection.Controls.Add(this.btnTestDbConnection);
            this.tabDbConnection.Location = new System.Drawing.Point(4, 22);
            this.tabDbConnection.Name = "tabDbConnection";
            this.tabDbConnection.Padding = new System.Windows.Forms.Padding(3);
            this.tabDbConnection.Size = new System.Drawing.Size(432, 400);
            this.tabDbConnection.TabIndex = 2;
            this.tabDbConnection.Text = "DB Connection";
            this.tabDbConnection.UseVisualStyleBackColor = true;
            // 
            // lblDbType
            // 
            this.lblDbType.AutoSize = true;
            this.lblDbType.Location = new System.Drawing.Point(184, 199);
            this.lblDbType.Name = "lblDbType";
            this.lblDbType.Size = new System.Drawing.Size(80, 13);
            this.lblDbType.TabIndex = 15;
            this.lblDbType.Text = "Database Type";
            // 
            // cbDatabaseType
            // 
            this.cbDatabaseType.FormattingEnabled = true;
            this.cbDatabaseType.Items.AddRange(new object[] {
            "ORACLE",
            "SQLSERVER"});
            this.cbDatabaseType.Location = new System.Drawing.Point(276, 194);
            this.cbDatabaseType.Name = "cbDatabaseType";
            this.cbDatabaseType.Size = new System.Drawing.Size(121, 21);
            this.cbDatabaseType.TabIndex = 14;
            this.cbDatabaseType.SelectedIndexChanged += new System.EventHandler(this.cbDatabaseType_SelectedIndexChanged);
            // 
            // tabcDbConnectionMode
            // 
            this.tabcDbConnectionMode.Controls.Add(this.tabConnectionString);
            this.tabcDbConnectionMode.Controls.Add(this.tabConnectionItem);
            this.tabcDbConnectionMode.Location = new System.Drawing.Point(7, 7);
            this.tabcDbConnectionMode.Name = "tabcDbConnectionMode";
            this.tabcDbConnectionMode.SelectedIndex = 0;
            this.tabcDbConnectionMode.Size = new System.Drawing.Size(390, 181);
            this.tabcDbConnectionMode.TabIndex = 13;
            // 
            // tabConnectionString
            // 
            this.tabConnectionString.Controls.Add(this.tbConnectionString);
            this.tabConnectionString.Controls.Add(this.lblConnectionString);
            this.tabConnectionString.Location = new System.Drawing.Point(4, 22);
            this.tabConnectionString.Name = "tabConnectionString";
            this.tabConnectionString.Padding = new System.Windows.Forms.Padding(3);
            this.tabConnectionString.Size = new System.Drawing.Size(382, 155);
            this.tabConnectionString.TabIndex = 0;
            this.tabConnectionString.Text = "Connection String";
            this.tabConnectionString.UseVisualStyleBackColor = true;
            // 
            // tbConnectionString
            // 
            this.tbConnectionString.Location = new System.Drawing.Point(10, 23);
            this.tbConnectionString.Multiline = true;
            this.tbConnectionString.Name = "tbConnectionString";
            this.tbConnectionString.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbConnectionString.Size = new System.Drawing.Size(366, 126);
            this.tbConnectionString.TabIndex = 1;
            // 
            // lblConnectionString
            // 
            this.lblConnectionString.AutoSize = true;
            this.lblConnectionString.Location = new System.Drawing.Point(7, 7);
            this.lblConnectionString.Name = "lblConnectionString";
            this.lblConnectionString.Size = new System.Drawing.Size(91, 13);
            this.lblConnectionString.TabIndex = 0;
            this.lblConnectionString.Text = "Connection String";
            // 
            // tabConnectionItem
            // 
            this.tabConnectionItem.Controls.Add(this.gbDatabaseConnection);
            this.tabConnectionItem.Location = new System.Drawing.Point(4, 22);
            this.tabConnectionItem.Name = "tabConnectionItem";
            this.tabConnectionItem.Padding = new System.Windows.Forms.Padding(3);
            this.tabConnectionItem.Size = new System.Drawing.Size(382, 155);
            this.tabConnectionItem.TabIndex = 1;
            this.tabConnectionItem.Text = "DB Item";
            this.tabConnectionItem.UseVisualStyleBackColor = true;
            // 
            // gbDatabaseConnection
            // 
            this.gbDatabaseConnection.Controls.Add(this.lblOraSource);
            this.gbDatabaseConnection.Controls.Add(this.cbOraSource);
            this.gbDatabaseConnection.Controls.Add(this.lblOraServers);
            this.gbDatabaseConnection.Controls.Add(this.cbOraServers);
            this.gbDatabaseConnection.Controls.Add(this.tbOraFilePath);
            this.gbDatabaseConnection.Controls.Add(this.lblOraFilePath);
            this.gbDatabaseConnection.Controls.Add(this.lblCipheredPassword);
            this.gbDatabaseConnection.Controls.Add(this.lblDbPassword);
            this.gbDatabaseConnection.Controls.Add(this.lblDbUser);
            this.gbDatabaseConnection.Controls.Add(this.tbDbCipheredPassword);
            this.gbDatabaseConnection.Controls.Add(this.tbDbPassword);
            this.gbDatabaseConnection.Controls.Add(this.tbDbUser);
            this.gbDatabaseConnection.Controls.Add(this.tbDbName);
            this.gbDatabaseConnection.Controls.Add(this.lblDbName);
            this.gbDatabaseConnection.Location = new System.Drawing.Point(-4, 7);
            this.gbDatabaseConnection.Name = "gbDatabaseConnection";
            this.gbDatabaseConnection.Size = new System.Drawing.Size(391, 142);
            this.gbDatabaseConnection.TabIndex = 1;
            this.gbDatabaseConnection.TabStop = false;
            this.gbDatabaseConnection.Text = "Database Connection";
            // 
            // lblOraSource
            // 
            this.lblOraSource.AutoSize = true;
            this.lblOraSource.Location = new System.Drawing.Point(200, 28);
            this.lblOraSource.Name = "lblOraSource";
            this.lblOraSource.Size = new System.Drawing.Size(67, 13);
            this.lblOraSource.TabIndex = 14;
            this.lblOraSource.Text = "ORA Source";
            this.lblOraSource.Visible = false;
            // 
            // cbOraSource
            // 
            this.cbOraSource.Enabled = false;
            this.cbOraSource.FormattingEnabled = true;
            this.cbOraSource.Location = new System.Drawing.Point(280, 25);
            this.cbOraSource.Name = "cbOraSource";
            this.cbOraSource.Size = new System.Drawing.Size(100, 21);
            this.cbOraSource.TabIndex = 13;
            this.cbOraSource.Visible = false;
            this.cbOraSource.SelectedIndexChanged += new System.EventHandler(this.cbOraSource_SelectedIndexChanged);
            // 
            // lblOraServers
            // 
            this.lblOraServers.AutoSize = true;
            this.lblOraServers.Location = new System.Drawing.Point(200, 81);
            this.lblOraServers.Name = "lblOraServers";
            this.lblOraServers.Size = new System.Drawing.Size(43, 13);
            this.lblOraServers.TabIndex = 12;
            this.lblOraServers.Text = "Servers";
            this.lblOraServers.Visible = false;
            // 
            // cbOraServers
            // 
            this.cbOraServers.Enabled = false;
            this.cbOraServers.FormattingEnabled = true;
            this.cbOraServers.Location = new System.Drawing.Point(280, 78);
            this.cbOraServers.Name = "cbOraServers";
            this.cbOraServers.Size = new System.Drawing.Size(100, 21);
            this.cbOraServers.TabIndex = 11;
            this.cbOraServers.Visible = false;
            this.cbOraServers.SelectedIndexChanged += new System.EventHandler(this.cbOraServers_SelectedIndexChanged);
            // 
            // tbOraFilePath
            // 
            this.tbOraFilePath.Enabled = false;
            this.tbOraFilePath.Location = new System.Drawing.Point(280, 52);
            this.tbOraFilePath.Name = "tbOraFilePath";
            this.tbOraFilePath.Size = new System.Drawing.Size(100, 20);
            this.tbOraFilePath.TabIndex = 10;
            this.tbOraFilePath.Visible = false;
            this.tbOraFilePath.Leave += new System.EventHandler(this.tbOraFilePath_Leave);
            // 
            // lblOraFilePath
            // 
            this.lblOraFilePath.AutoSize = true;
            this.lblOraFilePath.Location = new System.Drawing.Point(200, 55);
            this.lblOraFilePath.Name = "lblOraFilePath";
            this.lblOraFilePath.Size = new System.Drawing.Size(74, 13);
            this.lblOraFilePath.TabIndex = 9;
            this.lblOraFilePath.Text = "ORA File Path";
            this.lblOraFilePath.Visible = false;
            // 
            // lblCipheredPassword
            // 
            this.lblCipheredPassword.AutoSize = true;
            this.lblCipheredPassword.Location = new System.Drawing.Point(7, 104);
            this.lblCipheredPassword.Name = "lblCipheredPassword";
            this.lblCipheredPassword.Size = new System.Drawing.Size(75, 13);
            this.lblCipheredPassword.TabIndex = 8;
            this.lblCipheredPassword.Text = "Ciphered Pass";
            // 
            // lblDbPassword
            // 
            this.lblDbPassword.AutoSize = true;
            this.lblDbPassword.Location = new System.Drawing.Point(6, 81);
            this.lblDbPassword.Name = "lblDbPassword";
            this.lblDbPassword.Size = new System.Drawing.Size(53, 13);
            this.lblDbPassword.TabIndex = 7;
            this.lblDbPassword.Text = "Password";
            // 
            // lblDbUser
            // 
            this.lblDbUser.AutoSize = true;
            this.lblDbUser.Location = new System.Drawing.Point(6, 55);
            this.lblDbUser.Name = "lblDbUser";
            this.lblDbUser.Size = new System.Drawing.Size(47, 13);
            this.lblDbUser.TabIndex = 6;
            this.lblDbUser.Text = "DB User";
            // 
            // tbDbCipheredPassword
            // 
            this.tbDbCipheredPassword.Location = new System.Drawing.Point(88, 101);
            this.tbDbCipheredPassword.Name = "tbDbCipheredPassword";
            this.tbDbCipheredPassword.Size = new System.Drawing.Size(100, 20);
            this.tbDbCipheredPassword.TabIndex = 4;
            // 
            // tbDbPassword
            // 
            this.tbDbPassword.Location = new System.Drawing.Point(88, 75);
            this.tbDbPassword.Name = "tbDbPassword";
            this.tbDbPassword.Size = new System.Drawing.Size(100, 20);
            this.tbDbPassword.TabIndex = 3;
            // 
            // tbDbUser
            // 
            this.tbDbUser.Location = new System.Drawing.Point(88, 49);
            this.tbDbUser.Name = "tbDbUser";
            this.tbDbUser.Size = new System.Drawing.Size(100, 20);
            this.tbDbUser.TabIndex = 2;
            // 
            // tbDbName
            // 
            this.tbDbName.Location = new System.Drawing.Point(88, 23);
            this.tbDbName.Name = "tbDbName";
            this.tbDbName.Size = new System.Drawing.Size(100, 20);
            this.tbDbName.TabIndex = 1;
            // 
            // lblDbName
            // 
            this.lblDbName.AutoSize = true;
            this.lblDbName.Location = new System.Drawing.Point(6, 26);
            this.lblDbName.Name = "lblDbName";
            this.lblDbName.Size = new System.Drawing.Size(53, 13);
            this.lblDbName.TabIndex = 0;
            this.lblDbName.Text = "DB Name";
            // 
            // gbQuery
            // 
            this.gbQuery.Controls.Add(this.btnExecuteQuery);
            this.gbQuery.Controls.Add(this.tbDbQuery);
            this.gbQuery.Location = new System.Drawing.Point(6, 223);
            this.gbQuery.Name = "gbQuery";
            this.gbQuery.Size = new System.Drawing.Size(391, 133);
            this.gbQuery.TabIndex = 12;
            this.gbQuery.TabStop = false;
            this.gbQuery.Text = "Query";
            // 
            // btnExecuteQuery
            // 
            this.btnExecuteQuery.Location = new System.Drawing.Point(9, 86);
            this.btnExecuteQuery.Name = "btnExecuteQuery";
            this.btnExecuteQuery.Size = new System.Drawing.Size(126, 23);
            this.btnExecuteQuery.TabIndex = 11;
            this.btnExecuteQuery.Text = "Execute &Query";
            this.btnExecuteQuery.UseVisualStyleBackColor = true;
            this.btnExecuteQuery.Click += new System.EventHandler(this.btnExecuteQuery_Click);
            // 
            // tbDbQuery
            // 
            this.tbDbQuery.Location = new System.Drawing.Point(6, 19);
            this.tbDbQuery.Multiline = true;
            this.tbDbQuery.Name = "tbDbQuery";
            this.tbDbQuery.Size = new System.Drawing.Size(379, 60);
            this.tbDbQuery.TabIndex = 10;
            this.tbDbQuery.WordWrap = false;
            // 
            // btnTestDbConnection
            // 
            this.btnTestDbConnection.Location = new System.Drawing.Point(11, 194);
            this.btnTestDbConnection.Name = "btnTestDbConnection";
            this.btnTestDbConnection.Size = new System.Drawing.Size(167, 23);
            this.btnTestDbConnection.TabIndex = 10;
            this.btnTestDbConnection.Text = "&Test Connection";
            this.btnTestDbConnection.UseVisualStyleBackColor = true;
            this.btnTestDbConnection.Click += new System.EventHandler(this.btnTestDbConnection_Click_1);
            // 
            // tabEllipseSettings
            // 
            this.tabEllipseSettings.Controls.Add(this.drpEnvironment);
            this.tabEllipseSettings.Controls.Add(this.lblEnvironment);
            this.tabEllipseSettings.Controls.Add(this.btnStartEllipseSettings);
            this.tabEllipseSettings.Controls.Add(this.btnEllipseSettings);
            this.tabEllipseSettings.Controls.Add(this.btnEllipseAbout);
            this.tabEllipseSettings.Location = new System.Drawing.Point(4, 22);
            this.tabEllipseSettings.Name = "tabEllipseSettings";
            this.tabEllipseSettings.Size = new System.Drawing.Size(432, 400);
            this.tabEllipseSettings.TabIndex = 3;
            this.tabEllipseSettings.Text = "Ellipse Settings";
            this.tabEllipseSettings.UseVisualStyleBackColor = true;
            // 
            // drpEnvironment
            // 
            this.drpEnvironment.FormattingEnabled = true;
            this.drpEnvironment.Location = new System.Drawing.Point(6, 46);
            this.drpEnvironment.Name = "drpEnvironment";
            this.drpEnvironment.Size = new System.Drawing.Size(153, 21);
            this.drpEnvironment.TabIndex = 6;
            // 
            // lblEnvironment
            // 
            this.lblEnvironment.AutoSize = true;
            this.lblEnvironment.Location = new System.Drawing.Point(3, 29);
            this.lblEnvironment.Name = "lblEnvironment";
            this.lblEnvironment.Size = new System.Drawing.Size(69, 13);
            this.lblEnvironment.TabIndex = 5;
            this.lblEnvironment.Text = "Environment:";
            // 
            // btnStartEllipseSettings
            // 
            this.btnStartEllipseSettings.Location = new System.Drawing.Point(3, 3);
            this.btnStartEllipseSettings.Name = "btnStartEllipseSettings";
            this.btnStartEllipseSettings.Size = new System.Drawing.Size(156, 23);
            this.btnStartEllipseSettings.TabIndex = 3;
            this.btnStartEllipseSettings.Text = "Start &Ellipse Settings";
            this.btnStartEllipseSettings.UseVisualStyleBackColor = true;
            this.btnStartEllipseSettings.Click += new System.EventHandler(this.btnStartEllipseSettings_Click);
            // 
            // btnEllipseSettings
            // 
            this.btnEllipseSettings.Location = new System.Drawing.Point(84, 72);
            this.btnEllipseSettings.Name = "btnEllipseSettings";
            this.btnEllipseSettings.Size = new System.Drawing.Size(75, 23);
            this.btnEllipseSettings.TabIndex = 1;
            this.btnEllipseSettings.Text = "&Settings";
            this.btnEllipseSettings.UseVisualStyleBackColor = true;
            this.btnEllipseSettings.Click += new System.EventHandler(this.btnEllipseSettings_Click);
            // 
            // btnEllipseAbout
            // 
            this.btnEllipseAbout.Location = new System.Drawing.Point(3, 72);
            this.btnEllipseAbout.Name = "btnEllipseAbout";
            this.btnEllipseAbout.Size = new System.Drawing.Size(75, 23);
            this.btnEllipseAbout.TabIndex = 0;
            this.btnEllipseAbout.Text = "&About";
            this.btnEllipseAbout.UseVisualStyleBackColor = true;
            this.btnEllipseAbout.Click += new System.EventHandler(this.btnEllipseAbout_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 473);
            this.Controls.Add(this.tabcGeneral);
            this.Name = "MainForm";
            this.Text = "SharedClass Utility Software";
            this.tabcGeneral.ResumeLayout(false);
            this.tabHome.ResumeLayout(false);
            this.tabHome.PerformLayout();
            this.tabEncryption.ResumeLayout(false);
            this.tabEncryption.PerformLayout();
            this.tabDbConnection.ResumeLayout(false);
            this.tabDbConnection.PerformLayout();
            this.tabcDbConnectionMode.ResumeLayout(false);
            this.tabConnectionString.ResumeLayout(false);
            this.tabConnectionString.PerformLayout();
            this.tabConnectionItem.ResumeLayout(false);
            this.gbDatabaseConnection.ResumeLayout(false);
            this.gbDatabaseConnection.PerformLayout();
            this.gbQuery.ResumeLayout(false);
            this.gbQuery.PerformLayout();
            this.tabEllipseSettings.ResumeLayout(false);
            this.tabEllipseSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabcGeneral;
        private System.Windows.Forms.TabPage tabHome;
        private System.Windows.Forms.TabPage tabEncryption;
        private System.Windows.Forms.Button btnDecrypt;
        private System.Windows.Forms.TextBox tbCipherText;
        private System.Windows.Forms.Label lblPassPhrase;
        private System.Windows.Forms.Label lblEncryptText;
        private System.Windows.Forms.TextBox tbPassPhrase;
        private System.Windows.Forms.TextBox tbPlainText;
        private System.Windows.Forms.Button btnEncrypt;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Button btnPastePassPhrase;
        private System.Windows.Forms.Button btnPastePlainText;
        private System.Windows.Forms.Button btnCopyCipherText;
        private System.Windows.Forms.Button btnCleanPlainText;
        private System.Windows.Forms.Button btnCleanPassPhrase;
        private System.Windows.Forms.Label lblDeveloper;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblDevelopBy;
        private System.Windows.Forms.TabPage tabDbConnection;
        private System.Windows.Forms.TextBox tbDbQuery;
        private System.Windows.Forms.GroupBox gbQuery;
        private System.Windows.Forms.Button btnExecuteQuery;
        private System.Windows.Forms.Label lblDbType;
        private System.Windows.Forms.ComboBox cbDatabaseType;
        private System.Windows.Forms.TabControl tabcDbConnectionMode;
        private System.Windows.Forms.TabPage tabConnectionString;
        private System.Windows.Forms.TextBox tbConnectionString;
        private System.Windows.Forms.Label lblConnectionString;
        private System.Windows.Forms.TabPage tabConnectionItem;
        private System.Windows.Forms.GroupBox gbDatabaseConnection;
        private System.Windows.Forms.Label lblCipheredPassword;
        private System.Windows.Forms.Label lblDbPassword;
        private System.Windows.Forms.Label lblDbUser;
        private System.Windows.Forms.TextBox tbDbCipheredPassword;
        private System.Windows.Forms.TextBox tbDbPassword;
        private System.Windows.Forms.TextBox tbDbUser;
        private System.Windows.Forms.TextBox tbDbName;
        private System.Windows.Forms.Label lblDbName;
        private System.Windows.Forms.Button btnTestDbConnection;
        private System.Windows.Forms.Label lblOraServers;
        private System.Windows.Forms.ComboBox cbOraServers;
        private System.Windows.Forms.TextBox tbOraFilePath;
        private System.Windows.Forms.Label lblOraFilePath;
        private System.Windows.Forms.Label lblOraSource;
        private System.Windows.Forms.ComboBox cbOraSource;
        private System.Windows.Forms.TabPage tabEllipseSettings;
        private System.Windows.Forms.Button btnEllipseSettings;
        private System.Windows.Forms.Button btnEllipseAbout;
        private System.Windows.Forms.Button btnStartEllipseSettings;
        private System.Windows.Forms.Label lblEnvironment;
        private System.Windows.Forms.ComboBox drpEnvironment;
    }
}

