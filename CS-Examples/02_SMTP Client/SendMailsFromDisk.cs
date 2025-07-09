using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using OAuth;
using Spire.Email;
using Spire.Email.Smtp;

namespace SendMailsFromDisk
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

            //Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
            string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
            string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

            //scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.         
            string[] scopes = new string[]

            { "openid", "email", "offline_access","https://outlook.office.com/SMTP.Send"};

            //Creating AzureROPCTokenProvider object to get an access token from a token server.
            AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
            OAuthToken result = provider.GetAccessToken();

            //Creating SmtpClient object to connect Outlook server
            SmtpClient client = new SmtpClient(host, port, true, userName, result.Token, ConnectionProtocols.StartTls);
            client.Password = password;

            MailMessage msg = MailMessage.Load(this.filePath.Text, MailMessageFormat.Eml);

            client.SendOne(msg);
            MessageBox.Show("Sent Successfully");


            //Do not use OAuth2.0 authentication to connect to qq-email or Gmail
          /*  MailMessage msg = MailMessage.Load(this.filePath.Text, MailMessageFormat.Eml);
            SmtpClient client = new SmtpClient();
            client.Host = this.HostInput.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.ConnectionProtocols = ConnectionProtocols.Ssl;

            try
            {
                client.SendOne(msg);
                MessageBox.Show("Sent Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }*/

        }
        private void LoadBtn_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "Mail files|*.eml";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.filePath.Text = this.openFileDialog1.FileName;

            }
        }
    }
    
}