Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Smtp

Namespace SendMailsFromDisk
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim msg As MailMessage = MailMessage.Load(Me.filePath.Text,MailMessageFormat.Eml)
			Dim client As New SmtpClient()
			client.Host = Me.HostInput.Text
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl
			client.SendOne(msg)
		End Sub
		Private Sub LoadBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles LoadBtn.Click
            Me.openFileDialog1.Filter = "Mail files|*.eml"
            If Me.openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Me.filePath.Text = Me.openFileDialog1.FileName

            End If
		End Sub
	End Class

End Namespace