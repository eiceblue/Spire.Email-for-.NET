Imports System.ComponentModel
Imports System.Text
Imports OAuth
Imports Spire.Email
Imports Spire.Email.Smtp

Namespace SendBulkEmails
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private msgs As New List(Of MailMessage)()
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

			'scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.         
			Dim scopes() As String = { "openid", "email", "offline_access","https://outlook.office.com/SMTP.Send"}

			'Creating AzureROPCTokenProvider object to get an access token from a token server.
			Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
			Dim result As OAuthToken = provider.GetAccessToken()

			'Creating SmtpClient object to connect Outlook server
			Dim client As New SmtpClient(host, port, True, userName, result.Token, ConnectionProtocols.StartTls)
			client.Password = password

			'Send a email
			For i As Integer = 0 To msgs.Count - 1
				msgs(i).Cc.Add(Me.CcInput.Text)
				msgs(i).Bcc.Add(Me.BccInput.Text)
				msgs(i).Subject = Me.SubjectInput.Text
				msgs(i).BodyText = Me.TextInput.Text
			Next i
			Try
				client.SendSome(msgs)
				MessageBox.Show("Connected Successfully")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try



			'Do not use OAuth2.0 authentication to connect to qq-email or Gmail
'              SmtpClient client = new SmtpClient(this.HostInput.Text, int.Parse(this.PortInput.Text));
'              client.Username = this.Username.Text;
'              client.Password = this.Password.Text;
'              client.ConnectionProtocols = ConnectionProtocols.Ssl;
'              for (int i = 0; i < msgs.Count;i++ )
'              {
'                  msgs[i].Cc.Add(this.CcInput.Text);
'                  msgs[i].Bcc.Add(this.BccInput.Text);
'                  msgs[i].Subject = this.SubjectInput.Text;
'                  msgs[i].BodyText = this.TextInput.Text;
'              }
'              try
'              {
'                 client.SendSome(msgs);
'                 MessageBox.Show("Sent Successfully");
'              }
'              catch (Exception ex)
'              {
'                  MessageBox.Show(ex.Message);
'              }

		End Sub

		Private Sub AddReceiver_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AddReceiver.Click
			Dim mail As New MailMessage(New MailAddress(Me.FromAddress.Text, Me.DisplayName.Text), Me.ToAddress.Text)
			msgs.Add(mail)
			MessageBox.Show("Added")
			Me.ToAddress.Text = ""
		End Sub
	End Class
End Namespace