Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.IMap
Imports System.IO

Namespace GetFoldersInfo
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
			Dim folders As ImapFolderCollection = client.GetFolderCollection()
			Dim sb As New StringBuilder()
			For Each folderInfo As ImapFolder In folders
				' Folder name and get messages in the folder
				sb.AppendLine("Folder name is " & folderInfo.Name)
                sb.AppendLine("Message count: " & client.GetMessageCount(folderInfo.Name))
				sb.AppendLine("Is it Marked? " & folderInfo.Marked)
				sb.AppendLine("Message that recent flag count: " & folderInfo.RecentCount)
				sb.AppendLine("Is it Selectable? " & folderInfo.Selectable)
				sb.AppendLine("SubFolders count: " & folderInfo.SubFolders.Count)
				sb.AppendLine("The next unique identifier value: " & folderInfo.UidNext)
				sb.AppendLine("The unique identifier validity value: " & folderInfo.UidValidity)
				sb.AppendLine("Messages which are not set the seen flag count: " & folderInfo.UnseenCount)
				sb.AppendLine("-----------------------Next Folder---------------------------")
			Next folderInfo
            File.WriteAllText("GetFoldersInfo.txt", sb.ToString())
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace