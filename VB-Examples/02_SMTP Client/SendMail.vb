Imports OAuth
Imports Spire.Email
Imports Spire.Email.Smtp

Namespace SendMail
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			If rBSync.Checked Then
				SendEmailSynchronously()
			ElseIf rBAsync.Checked Then
				SendEmailAsynchronously()
			End If
		End Sub
		Private Function GetSmtpClient() As SmtpClient
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

			Return client


			'Do not use OAuth2.0 authentication to connect to qq-email or Gmail
'            SmtpClient client = new SmtpClient();
'            client.Host = this.HostInput.Text;
'            client.Username = this.Username.Text ;
'            client.Password = this.Password.Text ;
'            client.Port = int.Parse(this.PortInput.Text);
'            client.ConnectionProtocols = ConnectionProtocols.Ssl;
'            return client;
		End Function
		Private Sub SendEmailSynchronously()
			Dim msg As New MailMessage(Me.FromAddress.Text,Me.ToAddress.Text)
			setContent(msg)
			Dim client As SmtpClient = GetSmtpClient()
			Try
				client.SendOne(msg)
				MessageBox.Show("Sent Successfully")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

		End Sub
		Private Sub SendEmailAsynchronously()
			Dim msg As New MailMessage(New MailAddress(Me.FromAddress.Text,Me.DisplayName.Text),Me.ToAddress.Text)
			setContent(msg)
			Dim client As SmtpClient = GetSmtpClient()
			client.BeginSendOne(msg, Callback)

		End Sub
		Private Sub setContent(ByVal msg As MailMessage)
			msg.Cc.Add(Me.CcInput.Text)
			msg.Bcc.Add(Me.BccInput.Text)
			msg.Subject = Me.SubjectInput.Text
			msg.BodyText = Me.TextInput.Text
		End Sub
		Dim Callback As AsyncCallback = New AsyncCallback(AddressOf A)
		Private Sub A(ByVal ar As IAsyncResult)
			Dim task As IAsyncResultExt = TryCast(ar, IAsyncResultExt)
			If task IsNot Nothing AndAlso task.IsCanceled Then
				MessageBox.Show("Cancel")
			End If
			If task IsNot Nothing AndAlso task.errorInfor IsNot Nothing Then
				MessageBox.Show("{0}", task.errorInfor.Message)
			Else
				MessageBox.Show("Sent Successfully")
			End If
		End Sub
		Private Interface IAsyncResultExt
			Inherits IAsyncResult
			ReadOnly Property errorInfor() As Exception
			ReadOnly Property IsCanceled() As Boolean
		End Interface
	End Class
End Namespace