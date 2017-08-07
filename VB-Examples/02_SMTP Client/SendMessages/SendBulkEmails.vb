Imports System.ComponentModel
Imports System.Text
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
			Dim client As New SmtpClient(Me.HostInput.Text, Integer.Parse(Me.PortInput.Text))
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl
			For i As Integer = 0 To msgs.Count - 1
				msgs(i).Cc.Add(Me.CcInput.Text)
				msgs(i).Bcc.Add(Me.BccInput.Text)
				msgs(i).Subject = Me.SubjectInput.Text
				msgs(i).BodyText = Me.TextInput.Text
			Next i
			Try
			   client.SendSome(msgs)
			   MessageBox.Show("Sent Successfully")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

		End Sub

		Private Sub AddReceiver_Click(ByVal sender As Object, ByVal e As EventArgs) Handles AddReceiver.Click
			Dim mail As New MailMessage(New MailAddress(Me.FromAddress.Text, Me.DisplayName.Text), Me.ToAddress.Text)
			msgs.Add(mail)
			MessageBox.Show("Added")
			Me.ToAddress.Text = ""
		End Sub
	End Class
End Namespace