using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Outlook;

namespace AddMessage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Load PST file
            Spire.Email.Outlook.OutlookFile olf = new OutlookFile(@"..\..\..\..\..\..\Data\Sample.pst");
            //Load Outlook MSG file
            OutlookItem item = new OutlookItem();
            item.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.msg");
            //Select the "Inbox" folder
            OutlookFolder inboxFolder = olf.RootOutlookFolder.GetSubFolder("Inbox");
            //Add the MSG to "Inbox" folder
            inboxFolder.AddItem(item);
            MessageBox.Show("Completed");
        }
    }
}