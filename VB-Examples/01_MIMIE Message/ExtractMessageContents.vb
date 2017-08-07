Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports System.IO

Namespace ExtractMessageContents
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\Data\Sample.eml")
			Dim sb As New StringBuilder()
			sb.AppendLine("From:")
			sb.AppendLine(mail.From.Address)
			sb.AppendLine("To:")
			For Each toAddress As MailAddress In mail.To
				sb.AppendLine(toAddress.Address)
			Next toAddress
			sb.AppendLine("Subject:")
			sb.AppendLine(mail.Subject)
			If mail.BodyHtml IsNot Nothing Then
				sb.AppendLine(mail.BodyHtml)
			ElseIf mail.BodyText IsNot Nothing Then
				sb.AppendLine(mail.BodyText)
			End If
			File.WriteAllText("ExtractedContents.txt", sb.ToString())
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace