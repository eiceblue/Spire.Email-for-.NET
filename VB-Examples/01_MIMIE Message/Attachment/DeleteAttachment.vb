Imports System.ComponentModel
Imports System.Text
Imports Spire.Email

Namespace DeleteAttachment
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create an instance of MailMessage and load an email file
			Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\..\Data\AttachmentSample.eml")
			' Delete the attachment by index
			mail.Attachments.RemoveAt(1)
			' Delete the attachment by attachment name
			Dim i As Integer = 0
			Do While i < mail.Attachments.Count
				Dim attach As Attachment = mail.Attachments(i)
				If attach.ContentType.Name = "Sample.pdf" Then
					mail.Attachments.Remove(attach)
				End If
				i += 1
			Loop
			mail.Save("HasDeletedAttachment.eml", MailMessageFormat.Eml)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace