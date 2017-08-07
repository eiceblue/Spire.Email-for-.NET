using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;

namespace AddAttachments
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
            // Load an attachment
            Attachment attachment = new Attachment(@"..\..\..\..\..\..\..\Data\Sample.zip");
            // Add Multiple Attachment in instance of MailMessage class and Save message to disk
            mail.Attachments.Add(attachment);
            mail.Attachments.Add(new Attachment(@"..\..\..\..\..\..\..\Data\Sample.docx"));
            mail.Attachments.Add(new Attachment(@"..\..\..\..\..\..\..\Data\Sample.pdf"));
            mail.Save("AddAttachments.eml", MailMessageFormat.Eml);
            MessageBox.Show("Completed");
        }

    }

}