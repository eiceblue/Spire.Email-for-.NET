Imports System.ComponentModel
Imports System.Text
Imports Spire.Email.IMap

Namespace FilterMessagesFromMailbox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			' Create an imapclient with host, user and password
			Dim client As New ImapClient()
			client.Host = Me.HostInput.Text
			client.Username = Me.Username.Text
			client.Password = Me.Password.Text
			client.ConnectionProtocols = ConnectionProtocols.Ssl
			client.Port = Integer.Parse(Me.PortInput.Text)
			client.Connect()
			client.Select("Inbox")
			' Get a collection of messages
			Dim msgs As ImapMessageCollection = client.Search("'Subject' Contains 'Spire.Email'")
            MessageBox.Show("Imap: " & msgs.Count & " message(s) found.")
		End Sub
	End Class
End Namespace