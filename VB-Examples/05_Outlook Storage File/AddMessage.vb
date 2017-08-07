Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Outlook

Namespace AddMessage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load PST file
            Dim olf As Spire.Email.Outlook.OutlookFile = New OutlookFile("..\..\..\..\..\..\Data\Sample.pst")
			'Load Outlook MSG file
			Dim item As New OutlookItem()
            item.LoadFromFile("..\..\..\..\..\..\Data\Sample.msg")
			'Select the "Inbox" folder
			Dim inboxFolder As OutlookFolder = olf.RootOutlookFolder.GetSubFolder("Inbox")
			'Add the MSG to "Inbox" folder
            inboxFolder.AddItem(item)
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace