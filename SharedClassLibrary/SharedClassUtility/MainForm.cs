using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SharedClassLibrary.Utilities.Encryption;

namespace SharedClassUtility
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
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
            var text = tbText.Text.Trim();
            var passPhrase = string.IsNullOrWhiteSpace(tbPassPhrase.Text) ? null : tbPassPhrase.Text.Trim();

            var result = Encryption.Encrypt(text, passPhrase);
            tbResult.Text = result;
        }

        public void DecryptText()
        {
            var text = tbText.Text.Trim();
            var passPhrase = string.IsNullOrWhiteSpace(tbPassPhrase.Text) ? null : tbPassPhrase.Text.Trim();

            var result = Encryption.Decrypt(text, passPhrase);
            tbResult.Text = result;
        }

        private void btnGenerateKey_Click(object sender, EventArgs e)
        {
            //tbPassPhrase.Text = Encryption.GenerateKey();
        }
    }
}
