# spire.email csharp attachments
## add attachments to email message
```csharp
// Create a new email message
MailMessage mail = new MailMessage("From@domain.com", "To@domain.com");

// Create and add attachments
Attachment attachment = new Attachment(@"..\..\..\..\..\..\..\Data\Sample.zip");
mail.Attachments.Add(attachment);
mail.Attachments.Add(new Attachment(@"..\..\..\..\..\..\..\Data\Sample.docx"));
mail.Attachments.Add(new Attachment(@"..\..\..\..\..\..\..\Data\Sample.pdf"));
```

---

# Spire.Email Email Address Modification
## Change To and Cc email addresses in a MailMessage
```csharp
// A To address with a friendly name can also be specified
mail.To.Add(new MailAddress("To@domain.com", "e-iceblue"));
// Specify Cc email address along with a friendly name
mail.Cc.Add(new MailAddress("Cc@domain.com", "ExpectedName"));
```

---

# spire.email csharp email creation
## create email with embedded image
```csharp
//Set subject of the message
MailMessage mail = new MailMessage("From@domain.com", "To@domain.com");
//Add TO recipients
mail.To.Add("AddedTo@domain.com");
//Specify ReplyTo
mail.ReplyTo.Add("ReplyTo@domain.com");
//Add CC recipients
mail.Cc.Add("Cc@domain.com");
//Add BCC recipients
mail.Bcc.Add("Bcc@domain.com");
mail.Subject = "New message created by Spire.Email for .NET";
//Initialize linked resource for image
LinkedResource resource = new LinkedResource(@"..\..\..\..\..\..\Data\Background.png");
resource.ContentId = "iceblue.png";
//Add image to linked resource
mail.LinkedResources.Add(resource);
// Set body html
string htmlString = @"
<html>
<body background='cid:iceblue.png'>
<p> This is the Html body!</p>
</body>
</html>";
mail.BodyHtml = htmlString;
// Save message
mail.Save("CreateNewEmail.msg", MailMessageFormat.Msg);
```

---

# spire.email csharp email creation
## create and save email message with multiple recipients and formats
```csharp
// Set subject of the message
MailMessage mail = new MailMessage("From@domain.com", "To@domain.com");
// Add TO recipients
mail.To.Add("AddedTo@domain.com");
// specify ReplyTo 
mail.ReplyTo.Add("ReplyTo@domain.com");
// Add CC recipients
mail.Cc.Add("Cc@domain.com");
// Add BCC recipients
mail.Bcc.Add("Bcc@domain.com");
mail.Subject = "New message created by Spire.Email for .NET";
// Set Html body
mail.BodyHtml = "<html><body>This is the Html body</body></html>";
// Save message in EML, EMLX, MSG and MHTML formats
mail.Save("CreateNewEmail.msg", MailMessageFormat.Msg);
mail.Save("CreateNewEmail.eml", MailMessageFormat.Eml);
mail.Save("CreateNewEmail.emlx", MailMessageFormat.Emlx);
mail.Save("CreateNewEmail.mhtml", MailMessageFormat.Mhtml);
```

---

# Spire.Email C# Delete Attachment
## Delete email attachments by index or name
```csharp
// Create an instance of MailMessage and load an email file
MailMessage mail = MailMessage.Load("AttachmentSample.eml");

// Delete the attachment by index
mail.Attachments.RemoveAt(1);

// Delete the attachment by attachment name
for (int i = 0; i < mail.Attachments.Count; i++)
{
    Attachment attach = mail.Attachments[i];
    if (attach.ContentType.Name == "Sample.pdf")
    {
        mail.Attachments.Remove(attach);
    }
}

mail.Save("HasDeletedAttachment.eml", MailMessageFormat.Eml);
```

---

# spire.email csharp eml to mhtml
## Convert EML file to MHTML format
```csharp
MailMessage mail = MailMessage.Load("Sample.eml");
mail.Save("LoadAndSave.mhtml", MailMessageFormat.Mhtml);
```

---

# Spire.Email EML to MSG Conversion
## Convert EML email format to MSG format
```csharp
MailMessage mail = MailMessage.Load(@"..\..\..\..\..\..\..\Data\Sample.eml");
mail.Save("LoadAndSave.msg", MailMessageFormat.Msg);
```

---

# spire.email csharp attachment extraction
## extract attachments from email message
```csharp
// Create an instance of MailMessage and load an email file
MailMessage mail = MailMessage.Load("AttachmentSample.eml");
if (!Directory.Exists("Attachments"))
{
    Directory.CreateDirectory("Attachments");
}

foreach (Attachment attach in mail.Attachments)
{
    // To get and save the attachment
    string filePath = string.Format("Attachments\\{0}", attach.ContentType.Name);
    if (File.Exists(filePath))
    {
        File.Delete(filePath);
    }
    FileStream fs = File.Create(filePath);
    CopyStream(attach.Data, fs);
}

private void CopyStream(Stream input, Stream output) 
{
    byte[] buffer = new byte[8 * 1024];
    int len;
    while((len=input.Read(buffer,0,buffer.Length))>0)
    {
        output.Write(buffer, 0, len);
    }
}
```

---

# Spire.Email CSharp Email Content Extraction
## Extract contents from an email message including sender, recipients, subject, and body
```csharp
StringBuilder sb = new StringBuilder();
sb.AppendLine("From:");
sb.AppendLine(mail.From.Address);
sb.AppendLine("To:");
foreach (MailAddress toAddress in mail.To)
{
    sb.AppendLine(toAddress.Address);
}
sb.AppendLine("Subject:");
sb.AppendLine(mail.Subject);
if (mail.BodyHtml != null)
{
    sb.AppendLine(mail.BodyHtml);
}
else if (mail.BodyText != null)
{
    sb.AppendLine(mail.BodyText);
}
```

---

# Spire.Email Body Text Encoding
## Set body text encoding for email message
```csharp
// Create mail message
MailMessage mail = new MailMessage("From@domain.com", "To@domain.com");

// Set subject and body encoding
mail.Subject = "SetBodyTextEncoding";
mail.BodyTextCharset = Encoding.UTF8;
mail.BodyText = "This sample demo shows how to set body encoding";
```

---

# spire.email smtp client connection
## connect to SMTP server using OAuth2.0 authentication
```csharp
//Connect outlook-Office using OAuth2.0 authentication
string host = this.HostInput.Text;
string userName = this.Username.Text;
string password = this.Password.Text;
ushort port = ushort.Parse(this.PortInput.Text);

//Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

//scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.         
string[] scopes = new string[] { "openid", "email", "offline_access","https://outlook.office.com/SMTP.Send"};

//Creating AzureROPCTokenProvider object to get an access token from a token server.
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

//Creating SmtpClient object to connect Outlook server
SmtpClient client = new SmtpClient(host, port, true, userName, result.Token, ConnectionProtocols.StartTls);
client.Password = password;

//Send a email
Spire.Email.MailMessage msg = new Spire.Email.MailMessage(new Spire.Email.MailAddress(this.FromAddress.Text, this.DisplayName.Text), this.ToAddress.Text);
try
{
    client.SendOne(msg);
    MessageBox.Show("Connected Successfully");
}
catch (Exception ex)
{
    MessageBox.Show(ex.Message);
}
```

---

# spire.email csharp bulk email
## send bulk emails using OAuth2.0 authentication
```csharp
List<MailMessage> msgs = new List<MailMessage>();

// Connect outlook-Office using OAuth2.0 authentication
string host = this.HostInput.Text;
string userName = this.Username.Text;
string password = this.Password.Text;
ushort port = ushort.Parse(this.PortInput.Text);

// Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

// scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.
string[] scopes = new string[]
{ "openid", "email", "offline_access", "https://outlook.office.com/SMTP.Send" };

// Creating AzureROPCTokenProvider object to get an access token from a token server.
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Creating SmtpClient object to connect Outlook server
SmtpClient client = new SmtpClient(host, port, true, userName, result.Token, ConnectionProtocols.StartTls);
client.Password = password;

// Send a email
for (int i = 0; i < msgs.Count; i++)
{
    msgs[i].Cc.Add(this.CcInput.Text);
    msgs[i].Bcc.Add(this.BccInput.Text);
    msgs[i].Subject = this.SubjectInput.Text;
    msgs[i].BodyText = this.TextInput.Text;
}
try
{
    client.SendSome(msgs);
    MessageBox.Show("Connected Successfully");
}
catch (Exception ex)
{
    MessageBox.Show(ex.Message);
}

// Add receiver to the email list
private void AddReceiver_Click(object sender, EventArgs e)
{
    MailMessage mail = new MailMessage(new MailAddress(this.FromAddress.Text, this.DisplayName.Text), this.ToAddress.Text);
    msgs.Add(mail);
    MessageBox.Show("Added");
    this.ToAddress.Text = "";
}
```

---

# spire.email smtp client
## sending emails synchronously and asynchronously with oauth authentication
```csharp
private SmtpClient GetSmtpClient()
{
    string host = this.HostInput.Text;
    string userName = this.Username.Text;
    string password = this.Password.Text;
    ushort port = ushort.Parse(this.PortInput.Text);

    //Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
    string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
    string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

    //scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/SMTP.Send.         
    string[] scopes = new string[]
    { "openid", "email", "offline_access","https://outlook.office.com/SMTP.Send"};

    //Creating AzureROPCTokenProvider object to get an access token from a token server.
    AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
    OAuthToken result = provider.GetAccessToken();

    //Creating SmtpClient object to connect Outlook server
    SmtpClient client = new SmtpClient(host, port, true, userName, result.Token, ConnectionProtocols.StartTls);
    client.Password = password;

    return client;
}
private void SendEmailSynchronously()
{
    MailMessage msg = new MailMessage(this.FromAddress.Text,this.ToAddress.Text);
    setContent(msg);
    SmtpClient client = GetSmtpClient();
    client.SendOne(msg);
}
private void SendEmailAsynchronously()
{
    MailMessage msg = new MailMessage(new MailAddress(this.FromAddress.Text,this.DisplayName.Text),this.ToAddress.Text);
    setContent(msg);
    SmtpClient client = GetSmtpClient();
    client.BeginSendOne(msg, Callback);
}
private void setContent(MailMessage msg)
{
    msg.Cc.Add(this.CcInput.Text);
    msg.Bcc.Add(this.BccInput.Text);
    msg.Subject = this.SubjectInput.Text;
    msg.BodyText = this.TextInput.Text;
}
AsyncCallback Callback = delegate(IAsyncResult ar)
{
    IAsyncResultExt task = ar as IAsyncResultExt;
    if (task != null && task.IsCanceled)
    {
        // Handle cancellation
    }
    if (task != null && task.errorInfor != null)
    {
        // Handle error
    }
    else
    {
        // Handle success
    }
};
private interface IAsyncResultExt : IAsyncResult
{
    Exception errorInfor { get;}
    bool IsCanceled { get;}
}
```

---

# spire.email csharp smtp client
## send email from disk file using smtp client with oauth2.0 or basic authentication
```csharp
// OAuth2.0 authentication for Outlook/Office365
string host = "smtp.outlook.com";
string userName = "username@example.com";
string password = "password";
ushort port = 587;

// Register your application with Azure AD to use Microsoft/Office365 OAUTH
string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

// Define scopes
string[] scopes = new string[] { "openid", "email", "offline_access", "https://outlook.office.com/SMTP.Send" };

// Create AzureROPCTokenProvider to get an access token
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Create SmtpClient with OAuth2.0 authentication
SmtpClient client = new SmtpClient(host, port, true, userName, result.Token, ConnectionProtocols.StartTls);
client.Password = password;

// Load email message from disk
MailMessage msg = MailMessage.Load("email.eml", MailMessageFormat.Eml);

// Send the email
client.SendOne(msg);

// Alternative method without OAuth2.0 for other email providers
// Load email message from disk
MailMessage msg = MailMessage.Load("email.eml", MailMessageFormat.Eml);

// Create SmtpClient with basic authentication
SmtpClient client = new SmtpClient();
client.Host = "smtp.example.com";
client.Port = 465;
client.Username = "username@example.com";
client.Password = "password";
client.ConnectionProtocols = ConnectionProtocols.Ssl;

// Send the email
client.SendOne(msg);
```

---

# spire.email pop3 connection
## connect to POP3 server using OAuth2.0 or standard authentication
```csharp
// Connect outlook-Office using OAuth2.0 authentication
string host = "pop3_server_host";
string userName = "username";
string password = "password";
ushort port = 995;

// Register your application with Azure AD to use Microsoft/Office365 OAUTH
string tenantId = "tenant_id";
string clientId = "client_id";

// Scope: The scope to which the client is limited
string[] scopes = new string[] 
{ "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All" };

// Creating AzureROPCTokenProvider object to get an access token
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Creating Pop3Client object to connect Outlook server
Pop3Client client = new Pop3Client(host, port, userName, result.Token, true);
client.EnableSsl = true;
client.Connect();

// Standard connection without OAuth2.0 for other email services
Pop3Client clientStandard = new Pop3Client();
clientStandard.Host = "pop3_server_host";
clientStandard.Username = "username";
clientStandard.Password = "password";
clientStandard.Port = 995;
clientStandard.EnableSsl = true;
clientStandard.Connect();
```

---

# spire.email pop3 client with oauth
## get email from pop3 server using oauth authentication and save as eml file
```csharp
// Creating AzureROPCTokenProvider object to get an access token from a token server
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Creating Pop3Client object to connect Outlook server
Pop3Client client = new Pop3Client(host, port, userName, result.Token, true);
client.EnableSsl = true;

client.Connect();

// get Message by index and save
MailMessage msg = client.GetMessage(client.GetMessageCount());
msg.Save("GetMailAndSave.eml", MailMessageFormat.Eml);
```

---

# spire.email pop3 mailbox info
## retrieve mailbox size and message count using pop3 client with oauth2 authentication
```csharp
// Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

// scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/POP.AccessAsUser.All.
string[] scopes = new string[] { "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All" };

// Creating AzureROPCTokenProvider object to get an access token from a token server.
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Creating Pop3Client object to connect Outlook server
Pop3Client client = new Pop3Client(host, port, userName, result.Token, true);
client.EnableSsl = true;

client.Connect();

// Get the size of the mailbox, number of messages in the mailbox
long size = client.GetSize();
int msgCount = client.GetMessageCount();
```

---

# spire.email pop3 client
## connect to pop3 server using oauth and retrieve email information
```csharp
// Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application
string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

// scope: The scope to which the client is limited
string[] scopes = new string[] 
{ "openid", "email", "offline_access","https://outlook.office.com/POP.AccessAsUser.All"};

// Creating AzureROPCTokenProvider object to get an access token from a token server
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Creating Pop3Client object to connect Outlook server
Pop3Client client = new Pop3Client(host, port, userName, result.Token, true);
client.EnableSsl = true;

client.Connect();

Pop3MessageInfoCollection msgs = client.GetAllMessages();          
for (int i = 0; i < msgs.Count; i++)
{
    // Get message information
    string sequenceNumber = msgs[i].SequenceNumber.ToString();
    string size = msgs[i].Size.ToString();
    string uniqueId = msgs[i].UniqueId;
}
```

---

# spire.email pop3 client
## get email summary information using OAuth2.0 authentication
```csharp
// Register your application with Azure AD to use Microsoft/Office365 OAUTH in your application, and getting the tenantId and clientId from the registered application.
string tenantId = "9facb510-bbc6-4c95-a0f5-9a60d3315f3d";
string clientId = "0a7976ae-ca80-4cb8-983d-9fea41e88dac";

// scope: The scope to which the client is limited. Permissions for openid, offline_access and https://outlook.office.com/POP.AccessAsUser.All.         
string[] scopes = new string[] { "openid", "email", "offline_access","https://outlook.office.com/POP.AccessAsUser.All"};

// Creating AzureROPCTokenProvider object to get an access token from a token server.
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Creating Pop3Client object to connect Outlook server
Pop3Client client = new Pop3Client(host, port, userName, result.Token, true);
client.EnableSsl = true;
client.Connect();

for (int i = 1; i <= client.GetMessageCount(); i++)
{
    MailMessage msg = client.GetMessage(i);
    MessageBox.Show(string.Format("From:{0}\nTo:{1}\nSubject:{2}", msg.From, msg.To, msg.Subject));
}
```

---

# spire.email pop3 client
## get mail uid using pop3 client with oauth2.0 authentication
```csharp
// OAuth2.0 authentication setup
string tenantId = "tenant-id";
string clientId = "client-id";
string userName = "username";
string password = "password";
string[] scopes = new string[] { "openid", "email", "offline_access", "https://outlook.office.com/POP.AccessAsUser.All" };

// Get access token
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken result = provider.GetAccessToken();

// Connect to POP3 server with OAuth
Pop3Client client = new Pop3Client("host", 995, userName, result.Token, true);
client.EnableSsl = true;
client.Connect();

// Get message UID
string uid = client.GetMessagesUid(client.GetMessageCount());

// Alternative: Connect without OAuth2.0
Pop3Client client2 = new Pop3Client();
client2.Host = "host";
client2.Username = "username";
client2.Password = "password";
client2.Port = 995;
client2.EnableSsl = true;
client2.Connect();

// Get message UID
string uid2 = client2.GetMessagesUid(client2.GetMessageCount());
```

---

# Spire.Email IMAP Connection
## Connect to IMAP server using OAuth2 authentication

```csharp
// Define IMAP server connection parameters
string host = "imap.server.com";
ushort port = 993;
string userName = "username";
string password = "password";

// Define OAuth2 authentication parameters
string tenantId = "tenant-id";
string clientId = "client-id";

// Define the required scopes for OAuth2 authentication
string[] scopes = new string[]
{ "openid", "email", "offline_access", "https://outlook.office.com/IMAP.AccessAsUser.All"};

// Creating AzureROPCTokenProvider object to get an access token from a token server
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, userName, password, scopes);
OAuthToken x = provider.GetAccessToken();

// Creating ImapClient object to connect Outlook server
ImapClient client = new ImapClient(host, port, userName, x.Token, true, Spire.Email.IMap.ConnectionProtocols.Ssl);
client.Connect();
```

---

# spire.email imap folder operations
## create, delete and rename folders using IMAP client
```csharp
// Creating ImapClient object to connect Outlook server
ImapClient client = new ImapClient(host, port, userName, x.Token, true, Spire.Email.IMap.ConnectionProtocols.Ssl);
try
{
    client.Connect();
    MessageBox.Show("Connected Successfully");

    // Create a new folder
    client.CreateFolder("CreatedFolder");
    // Delete a folder
    client.DeleteFolder("Test");
    // Rename a folder
    client.RenameFolder("E-iceblue", "RenamedFolder");
    MessageBox.Show("Completed");
}
catch (Exception ex)
{
    MessageBox.Show(ex.Message);
}
```

---

# spire.email csharp imap filter
## Filter email messages from a mailbox using IMAP protocol with OAuth2 authentication
```csharp
// Create a token provider to get an access token
AzureROPCTokenProvider provider = new AzureROPCTokenProvider(tenantId, clientId, username, password, scopes);
OAuthToken token = provider.GetAccessToken();

// Create an IMAP client with OAuth authentication
ImapClient client = new ImapClient(host, port, username, token.Token, true, Spire.Email.IMap.ConnectionProtocols.Ssl);
client.Connect();

// Select the Inbox folder
client.Select("Inbox");

// Search for messages with subject containing specific text
ImapMessageCollection messages = client.Search("'Subject' Contains 'Spire.Email'");
```

---

# Spire.Email IMAP Client
## Get folder information from IMAP server
```csharp
// Creating ImapClient object to connect to IMAP server
ImapClient client = new ImapClient(host, port, userName, password, true, Spire.Email.IMap.ConnectionProtocols.Ssl);
client.Connect();

// Get folder collection from IMAP server
ImapFolderCollection folders = client.GetFolderCollection();

// Iterate through each folder and retrieve information
foreach (ImapFolder folderInfo in folders)
{
    // Folder name and get messages in the folder
    string folderName = folderInfo.Name;
    int messageCount = client.GetMessageCount(folderInfo.Name);
    bool isMarked = folderInfo.Marked;
    int recentCount = folderInfo.RecentCount;
    bool isSelectable = folderInfo.Selectable;
    int subFoldersCount = folderInfo.SubFolders.Count;
    string uidNext = folderInfo.UidNext;
    string uidValidity = folderInfo.UidValidity;
    int unseenCount = folderInfo.UnseenCount;
}
```

---

# spire.email csharp outlook storage
## add message to pst file
```csharp
//Load PST file
Spire.Email.Outlook.OutlookFile olf = new OutlookFile(@"..\..\..\..\..\..\Data\Sample.pst");
//Load Outlook MSG file
OutlookItem item = new OutlookItem();
item.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.msg");
//Select the "Inbox" folder
OutlookFolder inboxFolder = olf.RootOutlookFolder.GetSubFolder("Inbox");
//Add the MSG to "Inbox" folder
inboxFolder.AddItem(item);
```

---

# spire.email outlook subfolder
## Add a subfolder to Outlook storage file
```csharp
Spire.Email.Outlook.OutlookFile olf = new OutlookFile(@"..\..\..\..\..\..\Data\Sample.pst");
olf.RootOutlookFolder.AddFolder("NewFolder", "Spire");
```

---

# spire.email outlook pst
## Get subfolder information from Outlook PST file
```csharp
// Load PST file
OutlookFile olf = new OutlookFile(@"..\..\..\..\..\..\Data\Sample.pst");
// Get the folders collection
OutlookFolderCollection folderCollection = olf.RootOutlookFolder.GetSubFolders();
// Display folder name and number of message
foreach(OutlookFolder folderinfo in folderCollection)
{
    // Get folder name
    string folderName = folderinfo.Name;
    // Get total items count
    int totalItems = folderinfo.ItemCount;
    // Get unread items count
    int unreadItems = folderinfo.UnreadItemCount;
}
```

---

