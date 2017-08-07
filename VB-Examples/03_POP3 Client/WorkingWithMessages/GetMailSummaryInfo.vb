Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Pop3

Namespace GetMailSummaryInfo
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
			For i As Integer = 1 To client.GetMessageCount()
				Dim msg As MailMessage = client.GetMessage(i)
                MessageBox.Show(String.Format("From:{0}\nTo:{1}\nSubject:{2}", msg.From, msg.To, msg.Subject))
			Next i
		End Sub
	End Class
End Namespace