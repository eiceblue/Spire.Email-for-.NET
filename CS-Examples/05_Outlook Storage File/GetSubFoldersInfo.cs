using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.Outlook;
using System.IO;

namespace GetSubFoldersInfo
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
            OutlookFile olf = new OutlookFile(@"..\..\..\..\..\..\Data\Sample.pst");
            //Get the folders collection
            OutlookFolderCollection folderCollection = olf.RootOutlookFolder.GetSubFolders();
            StringBuilder sb = new StringBuilder();
            //Display folder name and number of message
            foreach(OutlookFolder folderinfo in folderCollection)
            {
                sb.AppendLine("Folder:" + folderinfo.Name);
                sb.AppendLine("Total items:" + folderinfo.ItemCount);
                sb.AppendLine("Total unread item:" + folderinfo.UnreadItemCount);
                sb.AppendLine("------------------Next Folder--------------------");
            }
            File.WriteAllText("GetSubFoldersInfo.txt", sb.ToString());
            MessageBox.Show("Completed");
        }
    }
}