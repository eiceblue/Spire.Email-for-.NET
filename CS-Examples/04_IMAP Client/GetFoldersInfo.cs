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
using OAuth;

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

		    //Connect outlook-Office using OAuth2.0 authentication
            /*   
             *Please use TargetFrameworkVersion 4.5 or above to run this demo
           
             *Dependencies
             AzureROPCTokenProvider.cs
             AzureTokenResponse.cs
             OAuthToken.cs
             TokenType.cs
             .\packages\Newtonsoft.Json.13.0.3\lib\net40\Newtonsoft.Json.dll
             */
			 
            string host = this.HostInput.Text;
            string userName = this.Username.Text;
            string password = this.Password.Text;
            ushort port = ushort.Parse(this.PortInput.Text);

            //Please note that using OAuth2 to authenticate logins must use Framworks 4.5 or above.
            //Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
            string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
            string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

            //scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/IMAP.AccessAsUser.All.         
            string[] scopes = new string[]
            { "openid", "email", "offline_access", "https://outlook.office.com/IMAP.AccessAsUser.All"};

            //Creating AzureROPCTokenProvider object to get an access token from a token server.
            AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
            OAuthToken x = provider.GetAccessToken();

            //Creating ImapClient object to connect Outlook server
            ImapClient client = new ImapClient(host, port, userName, x.Token, true, Spire.Email.IMap.ConnectionProtocols.Ssl);
            try
            {
                client.Connect();
                MessageBox.Show("Connected Successfully");

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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //Do not use OAuth2.0 authentication to connect to qq-email or Gmail
            // Create an imapclient with host, user and password
            /*ImapClient client = new ImapClient();
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
            MessageBox.Show("Completed");*/
        }
    }
}