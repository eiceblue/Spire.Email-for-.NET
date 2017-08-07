using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Smtp;

namespace SendBulkEmails
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        List<MailMessage> msgs = new List<MailMessage>();
        private void btnRun_Click(object sender, EventArgs e)
        {
            SmtpClient client = new SmtpClient(this.HostInput.Text, int.Parse(this.PortInput.Text));
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl;
            for (int i = 0; i < msgs.Count;i++ )
            {
                msgs[i].Cc.Add(this.CcInput.Text);
                msgs[i].Bcc.Add(this.BccInput.Text);
                msgs[i].Subject = this.SubjectInput.Text;
                msgs[i].BodyText = this.TextInput.Text;
            }
            try
            {
               client.SendSome(msgs);
               MessageBox.Show("Sent Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
         
        }

        private void AddReceiver_Click(object sender, EventArgs e)
        {
            MailMessage mail = new MailMessage(new MailAddress(this.FromAddress.Text, this.DisplayName.Text), this.ToAddress.Text);
            msgs.Add(mail);
            MessageBox.Show("Added");
            this.ToAddress.Text = "";
        }
    }
}