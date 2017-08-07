Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.IMap

Namespace ConnectToImapServer
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
            client.Port = Integer.Parse(Me.PortInput.Text)
            client.ConnectionProtocols = ConnectionProtocols.Ssl
            client.Connect()
            MessageBox.Show("Connected Successfully")
        End Sub
    End Class
End Namespace