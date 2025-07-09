Imports System.ComponentModel
Imports System.Text
Imports OAuth
Imports Spire.Email
Imports Spire.Email.Pop3

Namespace GetMailBoxInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click

		   'Connect outlook-Office using OAuth2.0 authentication
'               
'             *Please use TargetFrameworkVersion 4.5 or above to run this demo
'           
'             *Dependencies
'             AzureROPCTokenProvider.cs
'             AzureTokenResponse.cs
'             OAuthToken.cs
'             TokenType.cs
'             .\packages\Newtonsoft.Json.13.0.3\lib\net40\Newtonsoft.Json.dll
'             

			Dim host As String = Me.HostInput.Text
			Dim userName As String = Me.Username.Text
			Dim password As String = Me.Password.Text
			Dim port As UShort = UShort.Parse(Me.PortInput.Text)

			'Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
			Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
			Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

			'scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/POP.AccessAsUser.All.         
			Dim scopes() As String = { "openid", "email", "offline_access","https://outlook.office.com/POP.AccessAsUser.All"}

			'Creating AzureROPCTokenProvider object to get an access token from a token server.
			Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
			Dim result As OAuthToken = provider.GetAccessToken()

			'Creating Pop3Client object to connect Outlook server
			Dim client As New Pop3Client(host, port, userName, result.Token, True)
			client.EnableSsl = True
			Try
				client.Connect()
				MessageBox.Show("Connected Successfully")

				' Get the size of the mailbox, number of messages in the mailbox
				Dim size As Long = client.GetSize()
				Dim msgCount As Integer = client.GetMessageCount()
				MessageBox.Show(String.Format("The size of the mailbox :{0}; Message Count:{1}", size.ToString(), msgCount.ToString()))
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try


			'Do not use OAuth2.0 authentication to connect to qq-email or Gmail
			' Create a POP3 client
'                Pop3Client client = new Pop3Client();
'                client.Host = this.HostInput.Text;
'                client.Username = this.Username.Text;
'                client.Password = this.Password.Text;
'                client.Port = int.Parse(this.PortInput.Text);
'                client.EnableSsl = true;
'                client.Connect();
'                // Get the size of the mailbox, number of messages in the mailbox
'                long size = client.GetSize();
'                int msgCount = client.GetMessageCount();
'                MessageBox.Show(string.Format("The size of the mailbox :{0}; Message Count:{1}", size.ToString(), msgCount.ToString()));
		End Sub
	End Class
End Namespace