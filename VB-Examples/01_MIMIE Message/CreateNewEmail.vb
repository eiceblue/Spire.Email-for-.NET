Imports System.ComponentModel
Imports System.Text
Imports Spire.Email

Namespace CreateNewEmail
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Set subject of the message
			Dim mail As New MailMessage("From@domain.com", "To@domain.com")
			' Add TO recipients
			mail.To.Add("AddedTo@domain.com")
			' specify ReplyTo 
			mail.ReplyTo.Add("ReplyTo@domain.com")
			' Add CC recipients
			mail.Cc.Add("Cc@domain.com")
			' Add BCC recipients
			mail.Bcc.Add("Bcc@domain.com")
			mail.Subject = "New message created by Spire.Email for .NET"
			' Set Html body
			mail.BodyHtml = "<html><body>This is the Html body</body></html>"
			' Save message in EML, EMLX, MSG and MHTML formats
			mail.Save("CreateNewEmail.msg", MailMessageFormat.Msg)
			mail.Save("CreateNewEmail.eml", MailMessageFormat.Eml)
			mail.Save("CreateNewEmail.emlx", MailMessageFormat.Emlx)
			mail.Save("CreateNewEmail.mhtml", MailMessageFormat.Mhtml)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace