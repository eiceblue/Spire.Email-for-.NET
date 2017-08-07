using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email.IMap;

namespace FilterMessagesFromMailbox
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
            client.ConnectionProtocols = ConnectionProtocols.Ssl;
            client.Port = int.Parse(this.PortInput.Text);
            client.Connect();
            client.Select("Inbox");
            // Get a collection of messages
            ImapMessageCollection msgs = client.Search("'Subject' Contains 'Spire.Email'");
            MessageBox.Show("Imap: " + msgs.Count + " message(s) found.");
        }
    }
}