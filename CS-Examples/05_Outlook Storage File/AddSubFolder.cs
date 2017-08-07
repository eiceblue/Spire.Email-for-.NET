using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Outlook;

namespace AddSubFolder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Spire.Email.Outlook.OutlookFile olf = new OutlookFile(@"..\..\..\..\..\..\Data\Sample.pst");
            olf.RootOutlookFolder.AddFolder("NewFolder", "Spire");
            MessageBox.Show("Completed");
        }
    }
}