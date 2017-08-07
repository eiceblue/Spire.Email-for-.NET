using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.IMap;

namespace ConnectToImapServer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient();
            client.Host = this.HostInput.Text;
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.ConnectionProtocols = ConnectionProtocols.Ssl;
            client.Connect();
            MessageBox.Show("Connected Successfully");
        }
    }
}