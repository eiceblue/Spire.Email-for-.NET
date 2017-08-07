Imports System.ComponentModel
Imports System.Text
Imports Spire.Email.Smtp
Imports Spire.Email.Pop3

Namespace ConnectToPop3Server
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Try
				Dim client As New Pop3Client()
				client.Host = Me.HostInput.Text
				client.Username = Me.Username.Text
				client.Password = Me.Password.Text
				client.Port = Integer.Parse(Me.PortInput.Text)
				client.EnableSsl = True
				client.Connect()
                MessageBox.Show("Connected Successfully")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try
		End Sub
	End Class
End Namespace