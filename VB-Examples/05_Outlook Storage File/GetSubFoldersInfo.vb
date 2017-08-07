Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Outlook
Imports System.IO

Namespace GetSubFoldersInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load PST file
            Dim olf As New OutlookFile("..\..\..\..\..\..\Data\Sample.msg")
			'Get the folders collection
			Dim folderCollection As OutlookFolderCollection = olf.RootOutlookFolder.GetSubFolders()
            Dim sb As New StringBuilder()
            'Display folder name and number of message
			For Each folderinfo As OutlookFolder In folderCollection
                sb.AppendLine("Folder:" & folderinfo.Name)
                sb.AppendLine("Total items:" & folderinfo.ItemCount)
                sb.AppendLine("Total unread item:" & folderinfo.UnreadItemCount)
                sb.AppendLine("------------------Next Folder--------------------")
            Next folderinfo
            File.WriteAllText("GetSubFoldersInfo.txt", sb.ToString)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace