Imports System.ComponentModel
Imports System.Text
Imports Spire.Email
Imports Spire.Email.Smtp

Namespace SendMail
    Partial Public Class Form1
        Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            If rBSync.Checked Then
                SendEmailSynchronously()
            ElseIf rBAsync.Checked Then
                SendEmailAsynchronously()
            End If
        End Sub
        Private Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient()
            client.Host = Me.HostInput.Text
            client.Username = Me.Username.Text
            client.Password = Me.Password.Text
            client.Port = Integer.Parse(Me.PortInput.Text)
            client.ConnectionProtocols = Spire.Email.IMap.ConnectionProtocols.Ssl
            Return client
        End Function
        Private Sub SendEmailSynchronously()
            Dim msg As New MailMessage(Me.FromAddress.Text, Me.ToAddress.Text)
            setContent(msg)
            Dim client As SmtpClient = GetSmtpClient()
            Try
                client.SendOne(msg)
                MessageBox.Show("Sent Successfully")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub
        Private Sub SendEmailAsynchronously()
            Dim msg As New MailMessage(New MailAddress(Me.FromAddress.Text, Me.DisplayName.Text), Me.ToAddress.Text)
            setContent(msg)
            Dim client As SmtpClient = GetSmtpClient()
            client.BeginSendOne(msg, Callback)
        End Sub
        Private Sub setContent(ByVal msg As MailMessage)
            msg.Cc.Add(Me.CcInput.Text)
            msg.Bcc.Add(Me.BccInput.Text)
            msg.Subject = Me.SubjectInput.Text
            msg.BodyText = Me.TextInput.Text
        End Sub
        Dim Callback As AsyncCallback = New AsyncCallback(AddressOf A)
        Private Sub A(ByVal ar As IAsyncResult)
            Dim task As IAsyncResultExt = TryCast(ar, IAsyncResultExt)
            If task IsNot Nothing AndAlso task.IsCanceled Then
                MessageBox.Show("Cancel")
            End If
            If task IsNot Nothing AndAlso task.errorInfor IsNot Nothing Then
                MessageBox.Show("{0}", task.errorInfor.Message)
            Else
                MessageBox.Show("Sent Successfully")
            End If
        End Sub
        Private Interface IAsyncResultExt
            Inherits IAsyncResult
            ReadOnly Property errorInfor() As Exception
            ReadOnly Property IsCanceled() As Boolean
        End Interface

    End Class
End Namespace