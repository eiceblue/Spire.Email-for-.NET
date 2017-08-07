using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Pop3;

namespace GetMailSummaryInfo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a POP3 client
            Pop3Client client = new Pop3Client();
            client.Host = this.HostInput.Text;
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.EnableSsl = true;
            client.Connect();
            for (int i = 1; i <= client.GetMessageCount(); i++)
            {
                MailMessage msg = client.GetMessage(i);
                MessageBox.Show(string.Format("From:{0}\nTo:{1}\nSubject:{2}", msg.From, msg.To, msg.Subject));
            }
        }
    }
}