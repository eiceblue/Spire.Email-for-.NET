using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;

namespace SetBodyTextEncoding
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Set subject of the message
            MailMessage mail = new MailMessage("From@domain.com", "To@domain.com");
            // Set the body encoding.
            mail.Subject = "SetBodyTextEncoding";
            mail.BodyTextCharset = Encoding.UTF8;
            mail.BodyText = "This sample demo shows how to set body encoding";
            mail.Save("SetDefaultTextEncoding.eml", MailMessageFormat.Eml);
            MessageBox.Show("Completed");
        }
    }
}