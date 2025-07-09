using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Email.Smtp;
using Spire.Email;
using System.Reflection.Emit;
using OAuth;

namespace ConnectToServer
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

            //Send a email
            Spire.Email.MailMessage msg = new Spire.Email.MailMessage(new Spire.Email.MailAddress(this.FromAddress.Text, this.DisplayName.Text), this.ToAddress.Text);
            try
            {
                client.SendOne(msg);
                MessageBox.Show("Connected Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //Do not use OAuth2.0 authentication to connect to qq-email or Gmail
            /*SmtpClient client = new SmtpClient();
            client.Host = this.HostInput.Text;
            client.Username = this.Username.Text;
            client.Password = this.Password.Text;
            client.Port = int.Parse(this.PortInput.Text);
            client.ConnectionProtocols = ConnectionProtocols.Ssl;
            MailMessage message = new MailMessage(new MailAddress(this.FromAddress.Text, this.DisplayName.Text), this.ToAddress.Text);
            try
            {
                client.SendOne(message);
                MessageBox.Show("Connected Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }*/

        }
    }
}