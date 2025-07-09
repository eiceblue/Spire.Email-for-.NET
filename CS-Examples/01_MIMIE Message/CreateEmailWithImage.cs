using System;
using System.Windows.Forms;
using Spire.Email;

namespace CreateEmailWithImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
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
            //Add image to linked resoorce
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
            MessageBox.Show("Completed");
        }
    }
}