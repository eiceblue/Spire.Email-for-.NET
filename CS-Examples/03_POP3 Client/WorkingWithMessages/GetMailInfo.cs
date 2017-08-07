using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Pop3;

namespace GetMailInfo
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
            Pop3MessageInfoCollection msgs = client.GetAllMessages();
            for (int i = 0; i < msgs.Count; i++)
            {
                MessageBox.Show(string.Format("SequenceNumber:{0};Size:{1};UniqueId:{2}", msgs[i].SequenceNumber, msgs[i].Size, msgs[i].UniqueId));
            }
        }
    }
}