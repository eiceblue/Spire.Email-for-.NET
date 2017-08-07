using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;

namespace ChangeEmailAddress
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            MailMessage mail = MailMessage.Load(@"..\..\..\..\..\..\Data\Sample.eml");
            // A To address with a friendly name can also be specified
            mail.To.Add(new MailAddress("To@domain.com", "e-iceblue"));
            // Specify Cc email address along with a friendly name
            mail.Cc.Add(new MailAddress("Cc@domain.com", "ExpectedName"));
            mail.Save("ChangeEmailAddress.eml", MailMessageFormat.Eml);
            MessageBox.Show("Completed");
        }
    }
}