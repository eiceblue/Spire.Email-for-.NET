Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Pop3

Namespace GetMailInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a POP3 client
			Dim client As New Pop3Client()
			client.Host = Me.HostInput.Text
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.EnableSsl = True
			client.Connect()
			Dim msgs As Pop3MessageInfoCollection = client.GetAllMessages()
			For i As Integer = 0 To msgs.Count - 1
                MessageBox.Show(String.Format("SequenceNumber:{0};Size:{1};UniqueId:{2}", msgs(i).SequenceNumber, msgs(i).Size, msgs(i).UniqueId))
			Next i
		End Sub
	End Class
End Namespace