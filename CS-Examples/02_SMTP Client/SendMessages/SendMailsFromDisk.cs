using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Smtp;

namespace SendMailsFromDisk
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            MailMessage msg = MailMessage.Load(this.filePath.Text,MailMessageFormat.Eml);
            SmtpClient client = new SmtpClient();
            client.Host = this.HostInput.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl;
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
        private void LoadBtn_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Mail files|*.eml";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.filePath.Text = this.openFileDialog1.FileName;

            }
        }
    }
    
}