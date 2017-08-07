Imports System.ComponentModel
Imports System.Text
Imports Spire.Email

Namespace AddAttachments
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Set subject of the message
			Dim mail As New MailMessage("From@domain.com", "To@domain.com")
			' Load an attachment
            Dim attachment As New Attachment("..\..\..\..\..\..\..\Data\Sample.zip")
			' Add Multiple Attachment in instance of MailMessage class and Save message to disk
			mail.Attachments.Add(attachment)
			mail.Attachments.Add(New Attachment("..\..\..\..\..\..\..\Data\Sample.docx"))
			mail.Attachments.Add(New Attachment("..\..\..\..\..\..\..\Data\Sample.pdf"))
			mail.Save("AddAttachments.eml", MailMessageFormat.Eml)
            MessageBox.Show("Completed")
		End Sub

    End Class

End Namespace