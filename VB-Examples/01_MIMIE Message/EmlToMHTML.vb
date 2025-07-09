Imports System.ComponentModel
Imports System.Text
Imports Spire.Email

Namespace EmlToMHTML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\..\Data\Sample.eml")
			mail.Save("LoadAndSave.mhtml", MailMessageFormat.Mhtml)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace