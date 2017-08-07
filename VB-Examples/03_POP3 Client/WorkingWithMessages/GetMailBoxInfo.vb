Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Pop3

Namespace GetMailBoxInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a POP3 client
			Dim client As New Pop3Client()
			client.Host = Me.HostInput.Text
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.EnableSsl = True
			client.Connect()
			' Get the size of the mailbox, number of messages in the mailbox
			Dim size As Long = client.GetSize()
			Dim msgCount As Integer = client.GetMessageCount()
            MessageBox.Show(String.Format("The size of the mailbox :{0} Message Count:{1}", size.ToString(), msgCount.ToString()))
		End Sub
	End Class
End Namespace