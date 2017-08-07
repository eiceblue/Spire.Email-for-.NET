using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email;
using Spire.Email.IMap;
using System.IO;

namespace GetFoldersInfo
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
            ImapFolderCollection folders = client.GetFolderCollection();
            StringBuilder sb = new StringBuilder();
            foreach (ImapFolder folderInfo in folders)
            {
                // Folder name and get messages in the folder
                sb.AppendLine("Folder name is " + folderInfo.Name);
                sb.AppendLine("Message count: " + client.GetMessageCount(folderInfo.Name));
                sb.AppendLine("Is it Marked? " + folderInfo.Marked);
                sb.AppendLine("Message that recent flag count: " + folderInfo.RecentCount);
                sb.AppendLine("Is it Selectable? " + folderInfo.Selectable);
                sb.AppendLine("SubFolders count: " + folderInfo.SubFolders.Count);
                sb.AppendLine("The next unique identifier value: " + folderInfo.UidNext);
                sb.AppendLine("The unique identifier validity value: " + folderInfo.UidValidity);
                sb.AppendLine("Messages which are not set the seen flag count: " + folderInfo.UnseenCount);
                sb.AppendLine("-----------------------Next Folder---------------------------");
            }
            File.WriteAllText("GetFoldersInfo.txt", sb.ToString());
            MessageBox.Show("Completed");
        }
    }
}