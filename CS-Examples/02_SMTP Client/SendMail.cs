using System;
using System.Windows.Forms;
using OAuth;
using Spire.Email;
using Spire.Email.Smtp;

namespace SendMail
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            if(rBSync.Checked)
            {
                SendEmailSynchronously();
            }
            else if(rBAsync.Checked)
            {
                SendEmailAsynchronously();
            }
        }
        private SmtpClient GetSmtpClient()
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

            return client;


            //Do not use OAuth2.0 authentication to connect to qq-email or Gmail
            /*SmtpClient client = new SmtpClient();
            client.Host = this.HostInput.Text;
            client.Username = this.Username.Text ;
            client.Password = this.Password.Text ;
            client.Port = int.Parse(this.PortInput.Text);
            client.ConnectionProtocols = ConnectionProtocols.Ssl;
            return client;*/
        }
        private void SendEmailSynchronously()
        {
            MailMessage msg = new MailMessage(this.FromAddress.Text,this.ToAddress.Text);
            setContent(msg);
            SmtpClient client = GetSmtpClient();
            try
            {
                client.SendOne(msg);
                MessageBox.Show("Sent Successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void SendEmailAsynchronously()
        {
            MailMessage msg = new MailMessage(new MailAddress(this.FromAddress.Text,this.DisplayName.Text),this.ToAddress.Text);
            setContent(msg);
            SmtpClient client = GetSmtpClient();
            client.BeginSendOne(msg, Callback);
            
        }
        private void setContent(MailMessage msg)
        {
            msg.Cc.Add(this.CcInput.Text);
            msg.Bcc.Add(this.BccInput.Text);
            msg.Subject = this.SubjectInput.Text;
            msg.BodyText = this.TextInput.Text;
        }
        AsyncCallback Callback = delegate(IAsyncResult ar)
        {
            IAsyncResultExt task = ar as IAsyncResultExt;
            if (task != null && task.IsCanceled)
            {
                MessageBox.Show("Cancel");
            }
            if (task != null && task.errorInfor != null)
            {
                MessageBox.Show("{0}", task.errorInfor.Message);
            }
            else
            {
                MessageBox.Show("Sent Successfully");
            }
        };
        private interface IAsyncResultExt : IAsyncResult
        {
            Exception errorInfor { get;}
            bool IsCanceled { get;}
        }
    }
}