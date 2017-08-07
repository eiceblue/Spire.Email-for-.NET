using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;

namespace DeleteAttachment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create an instance of MailMessage and load an email file
            MailMessage mail = MailMessage.Load(@"..\..\..\..\..\..\..\Data\AttachmentSample.eml");
            // Delete the attachment by index
            mail.Attachments.RemoveAt(1);
            // Delete the attachment by attachment name
            for (int i = 0; i < mail.Attachments.Count; i++)
            {
                Attachment attach = mail.Attachments[i];
                if (attach.ContentType.Name == "Sample.pdf")
                {
                    mail.Attachments.Remove(attach);
                }
            }
            mail.Save("HasDeletedAttachment.eml", MailMessageFormat.Eml);
            MessageBox.Show("Completed");
        }        
    }
}