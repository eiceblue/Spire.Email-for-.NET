Imports System.ComponentModel
Imports System.Text
Imports Spire.Email

Namespace SetBodyTextEncoding
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Set subject of the message
			Dim mail As New MailMessage("From@domain.com", "To@domain.com")
			' Set the body encoding.
			mail.Subject = "SetBodyTextEncoding"
            mail.BodyTextCharset = Encoding.UTF8
            mail.BodyText = "This sample demo shows how to set body encoding"
			mail.Save("SetDefaultTextEncoding.eml", MailMessageFormat.Eml)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace