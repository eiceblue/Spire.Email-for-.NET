using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Pop3;

namespace GetMailUID
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a POP3 client and connect.
            Pop3Client client = new Pop3Client();
            client.Host = this.HostInput.Text;
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.EnableSsl = true;
            client.Connect();
            // get Message Uid by index
            string s = client.GetMessagesUid(client.GetMessageCount());
            MessageBox.Show(String.Format("Uid:{0}",s));
        }
    }
}