using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;

namespace EmlToMSG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            MailMessage mail = MailMessage.Load(@"..\..\..\..\..\..\..\Data\Sample.eml");
            mail.Save("LoadAndSave.msg", MailMessageFormat.Msg);
            MessageBox.Show("Completed");
        }
    }
}