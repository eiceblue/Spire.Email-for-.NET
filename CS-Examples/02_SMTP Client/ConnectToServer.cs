using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email.Smtp;
using Spire.Email;

namespace ConnectToServer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            SmtpClient client = new SmtpClient();
            client.Host = this.HostInput.Text ;
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl;
            MailMessage message = new MailMessage(new MailAddress(this.FromAddress.Text,this.DisplayName.Text),this.ToAddress.Text);
            try
            {
                client.SendOne(message);
                MessageBox.Show("Connected Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
     
        }       
    }
}