using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;

namespace CreateNewEmail
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
            // Add TO recipients
            mail.To.Add("AddedTo@domain.com");
            // specify ReplyTo 
            mail.ReplyTo.Add("ReplyTo@domain.com");
            // Add CC recipients
            mail.Cc.Add("Cc@domain.com");
            // Add BCC recipients
            mail.Bcc.Add("Bcc@domain.com");
            mail.Subject = "New message created by Spire.Email for .NET";
            // Set Html body
            mail.BodyHtml = "<html><body>This is the Html body</body></html>";
            // Save message in EML, EMLX, MSG and MHTML formats
            mail.Save("CreateNewEmail.msg", MailMessageFormat.Msg);
            mail.Save("CreateNewEmail.eml", MailMessageFormat.Eml);
            mail.Save("CreateNewEmail.emlx", MailMessageFormat.Emlx);
            mail.Save("CreateNewEmail.mhtml", MailMessageFormat.Mhtml);
            MessageBox.Show("Completed");
        }
    }
}