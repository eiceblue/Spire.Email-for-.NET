Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.IMap

Namespace CreateAndRenameFolder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create an imapclient with host, user and password
			Dim client As New ImapClient()
			client.Host = Me.HostInput.Text
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
            client.ConnectionProtocols = ConnectionProtocols.Ssl
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.Connect()
			' Create a new folder
			client.CreateFolder("E-iceblue1")
			' Delete a folder
			client.DeleteFolder("Test")
			' Rename a folder
            client.RenameFolder("E-iceblue", "E-iceblue2")
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace