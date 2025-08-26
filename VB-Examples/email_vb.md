# Spire.Email Add Attachments
## This code demonstrates how to add multiple attachments to an email message using Spire.Email library
```vbnet
' Create a new mail message
Dim mail As New MailMessage("From@domain.com", "To@domain.com")

' Load an attachment
Dim attachment As New Attachment("..\..\..\..\..\..\..\Data\Sample.zip")

' Add multiple attachments to the mail message
mail.Attachments.Add(attachment)
mail.Attachments.Add(New Attachment("..\..\..\..\..\..\..\Data\Sample.docx"))
mail.Attachments.Add(New Attachment("..\..\..\..\..\..\..\Data\Sample.pdf"))
```

---

# Change Email Address in Mail Message
## Demonstrates how to modify To and Cc addresses in an email message
```vbnet
' Load email message from file
Dim mail As MailMessage = MailMessage.Load("Sample.eml")

' A To address with a friendly name can also be specified
mail.To.Add(New MailAddress("To@domain.com", "e-iceblue"))

' Specify Cc email address along with a friendly name
mail.Cc.Add(New MailAddress("Cc@domain.com", "ExpectedName"))

' Save the modified email message
mail.Save("ChangeEmailAddress.eml", MailMessageFormat.Eml)
```

---

# spire.email vb.net email creation
## create email with embedded image
```vbnet
'Set subject of the message
Dim mail As New MailMessage("From@domain.com", "To@domain.com")
'Add TO recipients
mail.To.Add("AddedTo@domain.com")
'Specify ReplyTo
mail.ReplyTo.Add("ReplyTo@domain.com")
'Add CC recipients
mail.Cc.Add("Cc@domain.com")
'Add BCC recipients
mail.Bcc.Add("Bcc@domain.com")
mail.Subject = "New message created by Spire.Email for .NET"
'Initialize linked resource for image
Dim resource As New LinkedResource("..\..\..\..\..\..\Data\Background.png")
resource.ContentId = "iceblue.png"
'Add image to linked resource
mail.LinkedResources.Add(resource)
' Set body html
Dim htmlString As String = "" & ControlChars.CrLf & "            <html>" & ControlChars.CrLf & "            <body background='cid:iceblue.png'>" & ControlChars.CrLf & "            <p> This is the Html body!</p>" & ControlChars.CrLf & "            </body>" & ControlChars.CrLf & "            </html>"
mail.BodyHtml = htmlString
```

---

# spire.email vb.net email creation
## create and configure a new email message with recipients, subject and body
```vbnet
' Create a new email message
Dim mail As New MailMessage("From@domain.com", "To@domain.com")

' Add TO recipients
mail.To.Add("AddedTo@domain.com")

' Specify ReplyTo
mail.ReplyTo.Add("ReplyTo@domain.com")

' Add CC recipients
mail.Cc.Add("Cc@domain.com")

' Add BCC recipients
mail.Bcc.Add("Bcc@domain.com")

' Set subject of the message
mail.Subject = "New message created by Spire.Email for .NET"

' Set Html body
mail.BodyHtml = "<html><body>This is the Html body</body></html>"
```

---

# spire.email vb.net attachment
## delete email attachments
```vbnet
' Delete the attachment by index
mail.Attachments.RemoveAt(1)
' Delete the attachment by attachment name
Dim i As Integer = 0
Do While i < mail.Attachments.Count
    Dim attach As Attachment = mail.Attachments(i)
    If attach.ContentType.Name = "Sample.pdf" Then
        mail.Attachments.Remove(attach)
    End If
    i += 1
Loop
```

---

# spire.email vb.net conversion
## convert EML to MHTML format
```vbnet
' Load an EML file
Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\..\Data\Sample.eml")

' Save the loaded message as MHTML format
mail.Save("LoadAndSave.mhtml", MailMessageFormat.Mhtml)
```

---

# EML to MSG Conversion
## Convert EML email format to MSG format using Spire.Email
```vbnet
Dim mail As MailMessage = MailMessage.Load("input.eml")
mail.Save("output.msg", MailMessageFormat.Msg)
```

---

# spire.email vb.net attachment extraction
## extract attachments from email message
```vbnet
' Create an instance of MailMessage and load an email file
Dim mail As MailMessage = MailMessage.Load("..\..\..\..\..\..\..\Data\AttachmentSample.eml")
If Not Directory.Exists("Attachments") Then
    Directory.CreateDirectory("Attachments")
End If

For Each attach As Attachment In mail.Attachments
    ' To get and save the attachment
    Dim filePath As String = String.Format("Attachments\{0}", attach.ContentType.Name)
    If File.Exists(filePath) Then
        File.Delete(filePath)
    End If
    Dim fs As FileStream = File.Create(filePath)
    CopyStream(attach.Data, fs)
Next attach

Private Sub CopyStream(ByVal input As Stream, ByVal output As Stream)
    Dim buffer(8 * 1024 - 1) As Byte
    Dim len As Integer
    len=input.Read(buffer,0,buffer.Length)
    Do While len>0
        output.Write(buffer, 0, len)
        len=input.Read(buffer,0,buffer.Length)
    Loop
End Sub
```

---

# Spire.Email VB.NET Email Content Extraction
## Extract contents from an email message including sender, recipients, subject, and body
```vbnet
' Create a string builder to store the extracted contents
Dim sb As New StringBuilder()

' Extract sender information
sb.AppendLine("From:")
sb.AppendLine(mail.From.Address)

' Extract recipient information
sb.AppendLine("To:")
For Each toAddress As MailAddress In mail.To
    sb.AppendLine(toAddress.Address)
Next toAddress

' Extract subject
sb.AppendLine("Subject:")
sb.AppendLine(mail.Subject)

' Extract body content (HTML or plain text)
If mail.BodyHtml IsNot Nothing Then
    sb.AppendLine(mail.BodyHtml)
ElseIf mail.BodyText IsNot Nothing Then
    sb.AppendLine(mail.BodyText)
End If
```

---

# spire.email vb.net encoding
## set body text encoding for email message
```vbnet
' Set subject of the message
Dim mail As New MailMessage("From@domain.com", "To@domain.com")
' Set the body encoding.
mail.Subject = "SetBodyTextEncoding"
mail.BodyTextCharset = Encoding.UTF8
mail.BodyText = "This sample demo shows how to set body encoding"
```

---

# spire.email vb.net smtp client
## connect to SMTP server using OAuth2.0 authentication
```vbnet
'Connect outlook-Office using OAuth2.0 authentication
'Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

'scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.         
Dim scopes() As String = {"openid", "email", "offline_access", "https://outlook.office.com/SMTP.Send"}

'Creating AzureROPCTokenProvider object to get an access token from a token server.
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

'Creating SmtpClient object to connect Outlook server
Dim client As New SmtpClient(host, port, True, userName, result.Token, ConnectionProtocols.StartTls)
client.Password = password

'Send a email
Dim msg As New Spire.Email.MailMessage(New Spire.Email.MailAddress(fromAddress, displayName), toAddress)
Try
    client.SendOne(msg)
Catch ex As Exception
    'Handle exception
End Try
```

---

# spire.email vb.net bulk email
## send bulk emails using OAuth2.0 authentication
```vbnet
' Connect outlook-Office using OAuth2.0 authentication
Dim host As String = Me.HostInput.Text
Dim userName As String = Me.Username.Text
Dim password As String = Me.Password.Text
Dim port As UShort = UShort.Parse(Me.PortInput.Text)

' Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.         
Dim scopes() As String = { "openid", "email", "offline_access","https://outlook.office.com/SMTP.Send"}

' Creating AzureROPCTokenProvider object to get an access token from a token server.
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Creating SmtpClient object to connect Outlook server
Dim client As New SmtpClient(host, port, True, userName, result.Token, ConnectionProtocols.StartTls)
client.Password = password

' Send an email
For i As Integer = 0 To msgs.Count - 1
    msgs(i).Cc.Add(Me.CcInput.Text)
    msgs(i).Bcc.Add(Me.BccInput.Text)
    msgs(i).Subject = Me.SubjectInput.Text
    msgs(i).BodyText = Me.TextInput.Text
Next i
client.SendSome(msgs)

' Add receiver
Dim mail As New MailMessage(New MailAddress(Me.FromAddress.Text, Me.DisplayName.Text), Me.ToAddress.Text)
msgs.Add(mail)
```

---

# Spire.Email VB.NET SMTP Client
## Send emails using SMTP with OAuth2.0 authentication

' Function to get SMTP client with OAuth2.0 authentication
Private Function GetSmtpClient() As SmtpClient
    ' Connect outlook-Office using OAuth2.0 authentication
    ' Please use TargetFrameworkVersion 4.5 or above to run this demo
    
    Dim host As String = "smtp_host"
    Dim userName As String = "user_name"
    Dim password As String = "password"
    Dim port As UShort = port_number

    ' Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application
    Dim tenantId As String = "tenant_id"
    Dim clientId As String = "client_id"

    ' Scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.
    Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/SMTP.Send"}

    ' Creating AzureROPCTokenProvider object to get an access token from a token server.
    Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
    Dim result As OAuthToken = provider.GetAccessToken()

    ' Creating SmtpClient object to connect Outlook server
    Dim client As New SmtpClient(host, port, True, userName, result.Token, ConnectionProtocols.StartTls)
    client.Password = password

    Return client
End Function

' Method to send email synchronously
Private Sub SendEmailSynchronously()
    Dim msg As New MailMessage("from_address", "to_address")
    SetEmailContent(msg)
    Dim client As SmtpClient = GetSmtpClient()
    client.SendOne(msg)
End Sub

' Method to send email asynchronously
Private Sub SendEmailAsynchronously()
    Dim msg As New MailMessage(New MailAddress("from_address", "display_name"), "to_address")
    SetEmailContent(msg)
    Dim client As SmtpClient = GetSmtpClient()
    client.BeginSendOne(msg, AsyncCallback)
End Sub

' Method to set email content
Private Sub SetEmailContent(ByVal msg As MailMessage)
    msg.Cc.Add("cc_address")
    msg.Bcc.Add("bcc_address")
    msg.Subject = "email_subject"
    msg.BodyText = "email_body"
End Sub

' Callback method for asynchronous email sending
Private Sub AsyncCallback(ByVal ar As IAsyncResult)
    Dim task As IAsyncResultExt = TryCast(ar, IAsyncResultExt)
    If task IsNot Nothing AndAlso task.IsCanceled Then
        ' Handle cancellation
    End If
    If task IsNot Nothing AndAlso task.errorInfor IsNot Nothing Then
        ' Handle error
    Else
        ' Handle success
    End If
End Sub

' Interface for async result
Private Interface IAsyncResultExt
    Inherits IAsyncResult
    ReadOnly Property errorInfor() As Exception
    ReadOnly Property IsCanceled() As Boolean
End Interface

---

# spire.email vb.net smtp oauth
## send email from disk using OAuth2.0 authentication
```vbnet
' Get SMTP server settings
Dim host As String = Me.HostInput.Text
Dim userName As String = Me.Username.Text
Dim password As String = Me.Password.Text
Dim port As UShort = UShort.Parse(Me.PortInput.Text)

' Register your application with Azure AD to use Microsoft/Office365 OAUTH
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' Define OAuth scopes
Dim scopes() As String = { "openid", "email", "offline_access","https://outlook.office.com/SMTP.Send"}

' Create AzureROPCTokenProvider object to get an access token
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Create SmtpClient object to connect to the server with OAuth authentication
Dim client As New SmtpClient(host, port, True, userName, result.Token, ConnectionProtocols.StartTls)
client.Password = password

' Load email message from disk
Dim msg As MailMessage = MailMessage.Load(Me.filePath.Text, MailMessageFormat.Eml)

' Send the email
client.SendOne(msg)
```

---

# spire.email pop3 connection
## connect to POP3 server using OAuth2.0 authentication
```vbnet
' Get connection parameters
Dim host As String = Me.HostInput.Text
Dim userName As String = Me.Username.Text
Dim password As String = Me.Password.Text
Dim port As UShort = UShort.Parse(Me.PortInput.Text)

' Register your application with Azure AD to use Microsoft/Office365 OAUTH
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' Define scope for client permissions
Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All" }

' Create token provider to get access token from token server
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Create Pop3Client object to connect to Outlook server
Dim client As New Pop3Client(host, port, userName, result.Token, True)
client.EnableSsl = True
Try
    client.Connect()
    MessageBox.Show("Connected Successfully")
Catch ex As Exception
    MessageBox.Show(ex.Message)
End Try
```

---

# Spire.Email POP3 Client with OAuth2.0
## Retrieve email from POP3 server using OAuth2.0 authentication and save to EML file
```vbnet
' Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/POP.AccessAsUser.All.
Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All" }

' Creating AzureROPCTokenProvider object to get an access token from a token server.
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Creating Pop3Client object to connect Outlook server
Dim client As New Pop3Client(host, port, userName, result.Token, True)
client.EnableSsl = True

' Connect to the server
client.Connect()

' Get Message by index and save
Dim msg As MailMessage = client.GetMessage(client.GetMessageCount())
msg.Save("GetMailAndSave.eml", MailMessageFormat.Eml)
```

---

# spire.email vb.net pop3 client
## get mailbox information using OAuth2.0 authentication
```vbnet
' Connect outlook-Office using OAuth2.0 authentication
' Please use TargetFrameworkVersion 4.5 or above to run this demo
' Dependencies: AzureROPCTokenProvider.cs, AzureTokenResponse.cs, OAuthToken.cs, TokenType.cs, Newtonsoft.Json.dll

Dim host As String = Me.HostInput.Text
Dim userName As String = Me.Username.Text
Dim password As String = Me.Password.Text
Dim port As UShort = UShort.Parse(Me.PortInput.Text)

' Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application
' Getting the tenantId and clientId from the registered application
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' Scope: The scope to which the client is limited
' Permissions for openid, offline_access and https://outlook.office.com/POP.AccessAsUser.All
Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All"}

' Creating AzureROPCTokenProvider object to get an access token from a token server
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Creating Pop3Client object to connect Outlook server
Dim client As New Pop3Client(host, port, userName, result.Token, True)
client.EnableSsl = True

' Connect to the server
client.Connect()

' Get the size of the mailbox and number of messages in the mailbox
Dim size As Long = client.GetSize()
Dim msgCount As Integer = client.GetMessageCount()
```

---

# spire.email vb.net pop3 client
## get email information using OAuth2.0 authentication
```vbnet
' Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/POP.AccessAsUser.All.         
Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All" }

' Creating AzureROPCTokenProvider object to get an access token from a token server.
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Creating Pop3Client object to connect Outlook server
Dim client As New Pop3Client(host, port, userName, result.Token, True)
client.EnableSsl = True

client.Connect()

Dim msgs As Pop3MessageInfoCollection = client.GetAllMessages()
For i As Integer = 0 To msgs.Count - 1
    ' Process message information
    Dim sequenceNumber As Integer = msgs(i).SequenceNumber
    Dim size As Integer = msgs(i).Size
    Dim uniqueId As String = msgs(i).UniqueId
Next i
```

---

# VB.NET Email Summary Retrieval with OAuth
## Retrieve email summary information using POP3 client with OAuth authentication
```vbnet
' Create OAuth token provider for authentication
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Create and configure POP3 client
Dim client As New Pop3Client(host, port, userName, result.Token, True)
client.EnableSsl = True

' Connect to the email server
client.Connect()

' Retrieve and process email messages
For i As Integer = 1 To client.GetMessageCount()
    Dim msg As MailMessage = client.GetMessage(i)
    ' Email summary information can be accessed via msg.From, msg.To, msg.Subject
Next i
```

---

# spire.email vb.net pop3 client
## get mail unique identifier (UID) using POP3 protocol
```vbnet
' Creating AzureROPCTokenProvider object to get an access token from a token server
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim result As OAuthToken = provider.GetAccessToken()

' Creating Pop3Client object to connect Outlook server
Dim client As New Pop3Client(host, port, userName, result.Token, True)
client.EnableSsl = True

Try
    client.Connect()
    
    ' get Message Uid by index
    Dim s As String = client.GetMessagesUid(client.GetMessageCount())
Catch ex As Exception
    ' Handle exception
End Try

' Alternative method without OAuth2.0 authentication for qq-email or Gmail
' Create a POP3 client and connect
' Dim client As New Pop3Client()
' client.Host = host
' client.Username = userName
' client.Password = password
' client.Port = port
' client.EnableSsl = True
' client.Connect()
' get Message Uid by index
' Dim s As String = client.GetMessagesUid(client.GetMessageCount())
```

---

# VB.NET IMAP Server Connection
## Connect to IMAP server using OAuth2.0 authentication
```vbnet
' Get connection parameters
Dim host As String = Me.HostInput.Text
Dim userName As String = Me.Username.Text
Dim password As String = Me.Password.Text
Dim port As UShort = UShort.Parse(Me.PortInput.Text)

' Azure AD configuration for OAuth2.0 authentication
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

' Define OAuth scopes
Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/IMAP.AccessAsUser.All"}

' Create token provider and get access token
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim x As OAuthToken = provider.GetAccessToken()

' Create IMAP client with OAuth authentication
Dim client As New ImapClient(host, port, userName, x.Token, True, Spire.Email.IMap.ConnectionProtocols.Ssl)

' Connect to the server
Try
    client.Connect()
    MessageBox.Show("Connected Successfully")
Catch ex As Exception
    MessageBox.Show(ex.Message)
End Try
```

---

# IMAP Folder Operations
## Create, delete, and rename folders using IMAP client
```vbnet
' Creating ImapClient object to connect Outlook server
Dim client As New ImapClient(host, port, userName, accessToken, True, Spire.Email.IMap.ConnectionProtocols.Ssl)
client.Connect()

' Create a new folder
client.CreateFolder("CreatedFolder")
' Delete a folder
client.DeleteFolder("Test")
' Rename a folder
client.RenameFolder("E-iceblue", "RenamedFolder")
```

---

# spire.email vb.net imap filter
## filter messages from mailbox using IMAP protocol
```vbnet
'Connect outlook-Office using OAuth2.0 authentication
'
' *Please use TargetFrameworkVersion 4.5 or above to run this demo
'
' *Dependencies
' AzureROPCTokenProvider.cs
' AzureTokenResponse.cs
' OAuthToken.cs
' TokenType.cs
' .\packages\Newtonsoft.Json.13.0.3\lib\net40\Newtonsoft.Json.dll

Dim host As String = Me.HostInput.Text
Dim userName As String = Me.Username.Text
Dim password As String = Me.Password.Text
Dim port As UShort = UShort.Parse(Me.PortInput.Text)

'Please note that using OAuth2 to authenticate logins must use Framworks 4.5 or above.
'Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
Dim tenantId As String = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d"
Dim clientId As String = "0a7976ae-ca80-4cb8-983d-9fea41e88dac"

'scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/IMAP.AccessAsUser.All.         
Dim scopes() As String = { "openid", "email", "offline_access", "https://outlook.office.com/IMAP.AccessAsUser.All"}

'Creating AzureROPCTokenProvider object to get an access token from a token server.
Dim provider As New AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes)
Dim x As OAuthToken = provider.GetAccessToken()

'Creating ImapClient object to connect Outlook server
Dim client As New ImapClient(host, port, userName, x.Token, True, Spire.Email.IMap.ConnectionProtocols.Ssl)
Try
    client.Connect()
    MessageBox.Show("Connected Successfully")

    client.Select("Inbox")
    ' Get a collection of messages
    Dim msgs As ImapMessageCollection = client.Search("'Subject' Contains 'Spire.Email'")
    MessageBox.Show("Imap: " & msgs.Count & " message(s) found.")
Catch ex As Exception
    MessageBox.Show(ex.Message)
End Try
```

---

# spire.email vb.net imap
## get folder information from imap server
```vbnet
' Creating ImapClient object to connect Outlook server
Dim client As New ImapClient(host, port, userName, x.Token, True, Spire.Email.IMap.ConnectionProtocols.Ssl)
client.Connect()

Dim folders As ImapFolderCollection = client.GetFolderCollection()
For Each folderInfo As ImapFolder In folders
    ' Folder name and get messages in the folder
    Dim folderName As String = folderInfo.Name
    Dim messageCount As Integer = client.GetMessageCount(folderInfo.Name)
    Dim isMarked As Boolean = folderInfo.Marked
    Dim recentCount As Integer = folderInfo.RecentCount
    Dim isSelectable As Boolean = folderInfo.Selectable
    Dim subFoldersCount As Integer = folderInfo.SubFolders.Count
    Dim uidNext As String = folderInfo.UidNext
    Dim uidValidity As String = folderInfo.UidValidity
    Dim unseenCount As Integer = folderInfo.UnseenCount
Next folderInfo
```

---

# spire.email vb.net outlook storage
## add message to outlook pst file
```vbnet
'Load PST file
Dim olf As Spire.Email.Outlook.OutlookFile = New OutlookFile("Sample.pst")

'Load Outlook MSG file
Dim item As New OutlookItem()
item.LoadFromFile("Sample.msg")

'Select the "Inbox" folder
Dim inboxFolder As OutlookFolder = olf.RootOutlookFolder.GetSubFolder("Inbox")

'Add the MSG to "Inbox" folder
inboxFolder.AddItem(item)
```

---

# Spire.Email VB.NET Outlook Storage
## Add subfolder to Outlook PST file
```vbnet
' Create an Outlook file object from the PST file
Dim olf As Spire.Email.Outlook.OutlookFile = New OutlookFile("..\..\..\..\..\..\Data\Sample.pst")

' Add a new subfolder to the root folder
olf.RootOutlookFolder.AddFolder("NewFolder", "Spire")
```

---

# Outlook Subfolder Information Extraction
## Retrieve information about subfolders in an Outlook storage file including folder names, item counts, and unread item counts
```vbnet
' Load PST file
Dim olf As New OutlookFile(filePath)
' Get the folders collection
Dim folderCollection As OutlookFolderCollection = olf.RootOutlookFolder.GetSubFolders()
' Process each folder
For Each folderinfo As OutlookFolder In folderCollection
    ' Get folder information
    Dim folderName As String = folderinfo.Name
    Dim totalItems As Integer = folderinfo.ItemCount
    Dim unreadItems As Integer = folderinfo.UnreadItemCount
Next folderinfo
```

---

