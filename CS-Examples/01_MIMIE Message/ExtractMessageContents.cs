using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using System.IO;

namespace ExtractMessageContents
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
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("From:");
            sb.AppendLine(mail.From.Address);
            sb.AppendLine("To:");
            foreach (MailAddress toAddress in mail.To)
            {
                sb.AppendLine(toAddress.Address);
            }
            sb.AppendLine("Subject:");
            sb.AppendLine(mail.Subject);
            if (mail.BodyHtml != null)
            {
                sb.AppendLine(mail.BodyHtml);
            }
            else if (mail.BodyText != null)
            {
                sb.AppendLine(mail.BodyText);
            }
            File.WriteAllText("ExtractedContents.txt", sb.ToString());
            MessageBox.Show("Completed");
        }
    }
}