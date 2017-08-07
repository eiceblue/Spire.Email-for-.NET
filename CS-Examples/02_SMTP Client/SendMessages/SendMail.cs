using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Smtp;

namespace SendMail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            if(rBSync.Checked)
            {
                SendEmailSynchronously();
            }
            else if(rBAsync.Checked)
            {
                SendEmailAsynchronously();
            }
        }
        private SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient();
            client.Host = this.HostInput.Text;
            client.Username = this.Username.Text ;
            client.Password = this.Password.Text ;
            client.Port = int.Parse(this.PortInput.Text);
            client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl;
            return client;
        }
        private void SendEmailSynchronously()
        {
            MailMessage msg = new MailMessage(this.FromAddress.Text,this.ToAddress.Text);
            setContent(msg);
            SmtpClient client = GetSmtpClient();
            try
            {
                client.SendOne(msg);
                MessageBox.Show("Sent Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void SendEmailAsynchronously()
        {
            MailMessage msg = new MailMessage(new MailAddress(this.FromAddress.Text,this.DisplayName.Text),this.ToAddress.Text);
            setContent(msg);
            SmtpClient client = GetSmtpClient();
            client.BeginSendOne(msg, Callback);
            
        }
        private void setContent(MailMessage msg)
        {
            msg.Cc.Add(this.CcInput.Text);
            msg.Bcc.Add(this.BccInput.Text);
            msg.Subject = this.SubjectInput.Text;
            msg.BodyText = this.TextInput.Text;
        }
        AsyncCallback Callback = delegate(IAsyncResult ar)
        {
            IAsyncResultExt task = ar as IAsyncResultExt;
            if (task != null && task.IsCanceled)
            {
                MessageBox.Show("Cancel");
            }
            if (task != null && task.errorInfor != null)
            {
                MessageBox.Show("{0}", task.errorInfor.Message);
            }
            else
            {
                MessageBox.Show("Sent Successfully");
            }
        };
        private interface IAsyncResultExt : IAsyncResult
        {
            Exception errorInfor { get;}
            bool IsCanceled { get;}
        }
    }
}