Imports System.ComponentModel
Imports System.Text
Imports Spire.Email.Smtp
Imports Spire.Email

Namespace ConnectToServer
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim client As New SmtpClient()
			client.Host = Me.HostInput.Text
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl
			Dim message As New MailMessage(New MailAddress(Me.FromAddress.Text,Me.DisplayName.Text),Me.ToAddress.Text)
			Try
				client.SendOne(message)
				MessageBox.Show("Connected Successfully")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

		End Sub
	End Class
End Namespace