Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Outlook

Namespace AddSubFolder
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            Dim olf As Spire.Email.Outlook.OutlookFile = New OutlookFile("..\..\..\..\..\..\Data\Sample.pst")
            olf.RootOutlookFolder.AddFolder("NewFolder", "Spire")
            MessageBox.Show("Completed")
		End Sub
	End Class
End Namespace