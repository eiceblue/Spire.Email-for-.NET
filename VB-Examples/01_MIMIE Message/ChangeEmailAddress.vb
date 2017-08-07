Imports System.ComponentModel
Imports System.Text
Imports Spire.Email

Namespace ChangeEmailAddress
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\Data\Sample.eml")
			' A To address with a friendly name can also be specified
			mail.To.Add(New MailAddress("To@domain.com", "e-iceblue"))
			' Specify Cc email address along with a friendly name
			mail.Cc.Add(New MailAddress("Cc@domain.com", "ExpectedName"))
			mail.Save("ChangeEmailAddress.eml", MailMessageFormat.Eml)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace