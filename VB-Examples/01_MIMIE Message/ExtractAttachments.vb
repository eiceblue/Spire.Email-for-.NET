Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports System.IO

Namespace ExtractAttachments
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create an instance of MailMessage and load an email file
			Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\..\Data\AttachmentSample.eml")
			If Not Directory.Exists("Attachments") Then
				Directory.CreateDirectory("Attachments")
			End If

			For Each attach As Attachment In mail.Attachments
				' To get and save the attachment
				Dim filePath As String = String.Format("Attachments\{0}", attach.ContentType.Name)
				If File.Exists(filePath) Then
					File.Delete(filePath)
				End If
				Dim fs As FileStream = File.Create(filePath)
				CopyStream(attach.Data, fs)
			Next attach
            MessageBox.Show("Completed")
		End Sub
		Private Sub CopyStream(ByVal input As Stream, ByVal output As Stream)
			Dim buffer(8 * 1024 - 1) As Byte
			Dim len As Integer
			len=input.Read(buffer,0,buffer.Length)
			Do While len>0
				output.Write(buffer, 0, len)
				len=input.Read(buffer,0,buffer.Length)
			Loop
		End Sub
	End Class
End Namespace