Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Pop3

Namespace GetMailAndSave
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a POP3 client and connect.
			Dim client As New Pop3Client()
			client.Host = Me.HostInput.Text
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.EnableSsl = True
			client.Connect()
			' get Message by index and save
			Dim msg As MailMessage = client.GetMessage(client.GetMessageCount())
            msg.Save("GetMailAndSave.eml", MailMessageFormat.Eml)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace