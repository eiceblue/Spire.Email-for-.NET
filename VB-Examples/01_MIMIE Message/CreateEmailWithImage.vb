Imports Spire.Email

Namespace CreateEmailWithImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Set subject of the message
			Dim mail As New MailMessage("From@domain.com", "To@domain.com")
			'Add TO recipients
			mail.To.Add("AddedTo@domain.com")
			'Specify ReplyTo
			mail.ReplyTo.Add("ReplyTo@domain.com")
			'Add CC recipients
			mail.Cc.Add("Cc@domain.com")
			'Add BCC recipients
			mail.Bcc.Add("Bcc@domain.com")
			mail.Subject = "New message created by Spire.Email for .NET"
			'Initialize linked resource for image
			Dim resource As New LinkedResource("..\..\..\..\..\..\Data\Background.png")
			resource.ContentId = "iceblue.png"
			'Add image to linked resoorce
			mail.LinkedResources.Add(resource)
			' Set body html
			Dim htmlString As String = "" & ControlChars.CrLf & "            <html>" & ControlChars.CrLf & "            <body background='cid:iceblue.png'>" & ControlChars.CrLf & "            <p> This is the Html body!</p>" & ControlChars.CrLf & "            </body>" & ControlChars.CrLf & "            </html>"
			mail.BodyHtml = htmlString
			' Save message
			mail.Save("CreateNewEmail.msg", MailMessageFormat.Msg)
			MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace