using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using System.IO;

namespace ExtractAttachments
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
            if (!Directory.Exists("Attachments"))
            {
                Directory.CreateDirectory("Attachments");
            }
           
            foreach (Attachment attach in mail.Attachments)
            {
                // To get and save the attachment
                string filePath = string.Format("Attachments\\{0}", attach.ContentType.Name);
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                FileStream fs = File.Create(filePath);
                CopyStream(attach.Data, fs);
            }
            MessageBox.Show("Completed");
        }
        private void CopyStream(Stream input, Stream output) 
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while((len=input.Read(buffer,0,buffer.Length))>0)
            {
                output.Write(buffer, 0, len);
            }
        }
    }
}