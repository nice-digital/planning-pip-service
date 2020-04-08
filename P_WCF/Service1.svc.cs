using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Net.Mail;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Configuration;

namespace P_WCF
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    public class Service1 : IService1
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        public CompositeType GetDataUsingDataContract(CompositeType composite)
        {
            if (composite == null)
            {
                throw new ArgumentNullException("composite");
            }
            if (composite.BoolValue)
            {
                composite.StringValue += "Suffix";
            }
            return composite;
        }

        public void SendEmail(string whichapp, int P_EmailID)
        {

            Logger.Debug("Calling P_WCF for environment " + whichapp + " and email " + P_EmailID.ToString());
            
            string ConnectionString = WebConfigurationManager.AppSettings["ConnectionString"];
            string MailServer = WebConfigurationManager.AppSettings["MailServer"];

            // Create Instance of Connection and Command Object
            SqlConnection myConnection = new SqlConnection(ConnectionString);
            SqlCommand myCommand = new SqlCommand("P_EmailsToSend_List", myConnection);

            // Mark the Command as a SPROC
            myCommand.CommandType = CommandType.StoredProcedure;
            myCommand.CommandTimeout = 6000;

            // Add Parameters to SPROC
            myCommand.Parameters.Add("@P_EmailID", SqlDbType.Int).Value = P_EmailID;

            DataSet ds = new DataSet();

            try
            {
                //create the DataAdapter & DataSet
                myConnection.Open();
                SqlDataAdapter da = new SqlDataAdapter(myCommand);

                //fill the DataSet using default values for DataTable names, etc.
                da.Fill(ds, "EmailDetails");
            }
            catch (Exception ex)
            {
                ds.Dispose();
                Logger.Error(ex);
                throw;
            }
            finally
            {
                myConnection.Close();
            }

            DataTable dtMessage = ds.Tables[0];
            DataTable dtAttachments = ds.Tables[1];
            DataTable dtRecipients = ds.Tables[2];
            DataTable dtDoNotContacts = ds.Tables[3];

            DataRow dr3 = dtDoNotContacts.Rows[0];
            int NumNoSend = int.Parse(dr3["NumNoSend"].ToString());
            string UserEmail = dr3["UserEmail"].ToString();
            string User = dr3["FullName"].ToString();

            DataRow dr = dtMessage.Rows[0];
            string Message = dr["Message"].ToString();
            string Subject = dr["Subject"].ToString().Replace('\r', ' ').Replace('\n', ' ');
            string ReplyToAddress = dr["ReplyToAddress"].ToString();
            string MessageShort = dr["MessageShort"].ToString();
            string Header = dr["Header"].ToString();
            string MessagePlain = dr["MessagePlain"].ToString();

            SmtpClient smtpClient = new SmtpClient();

            string EmailAddress;
            string textLogo;
            string Recipient;

            textLogo = "<P style=\"MARGIN-BOTTOM:0px; TEXT-ALIGN:right;\"><STRONG><FONT color=\"#FFFFFF\" size=\"5\" face=\"Frutiger, Verdana, Arial, Helvetica, Sans-Serif\" style=\"BACKGROUND-COLOR: #0072C6;\"><EM>NHS</EM></FONT></STRONG></P>";
            textLogo += "<P style=\"MARGIN-TOP:6px; TEXT-ALIGN:right;\"><FONT size=\"4\" face=\"Arial, Helvetica, Sans-Serif\"><STRONG><EM>National Institute for</EM><BR><EM>Health and Care Excellence</EM></STRONG></FONT></P>";

            //create a header and footer
            string MessageHeader;
            string MessageFooter;
            string ManchesterHeader = "<p style=\"font-size:11px; margin-top:1em;\">Level 1A, City Tower<br />Piccadilly Plaza<br />Manchester<br />M1 4BT<br /><br />Tel: 0300 323 0140<br />Fax: 0845 003 7784<br /><br />www.nice.org.uk<br /></p><p>&nbsp;</p>";
            string LondonHeader = "<p style=\"font-size:11px; margin-top:1em;\">1st Floor<br />10 Spring Gardens<br />London<br />SW1A 2BU<br /><br />Tel: 0300 323 0140<br />Fax: 0845 003 7784<br /><br />www.nice.org.uk</p><p>&nbsp;</p>";
            string todaysDate = DateTime.Now.ToShortDateString();

            MessageHeader = "<html>";
            MessageHeader += "<head>";
            MessageHeader += "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />";
            MessageHeader += "<title>Email from NICE</title>";
            MessageHeader += "<style type=\"text/css\">";
            MessageHeader += "html, body { margin:20px; border:0; padding:0; text-align:center; }";
            MessageHeader += "body {";
            MessageHeader += "font-family: arial, san-serif, Verdana;";
            MessageHeader += "font-size: 12px;";
            MessageHeader += "text-align: left;";
            MessageHeader += "}";
            MessageHeader += "</style>";
            MessageHeader += "</head>";
            MessageHeader += "<body>";
            MessageHeader += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"width:800px;background-color:white\">";
            if (Header.ToString() == "Manchester" || Header.ToString() == "London")
            {
                //put in date
                MessageHeader += "<tr>";
                MessageHeader += "<td><p>" + todaysDate + "</p></td>";
                MessageHeader += "</tr>";

                //put in text version of logo
                //MessageHeader += "<tr>";
                //MessageHeader += "<td style=\"text-align:right;\">";
                //MessageHeader += textLogo;
                //MessageHeader += "</td>";
                //MessageHeader += "</tr>";

                //logo and address
                MessageHeader += "<tr>";
                MessageHeader += "<td>";
                MessageHeader += "<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"width:800px;background-color:white;margin-top:1em;\">";
                MessageHeader += "<tr>";
                MessageHeader += "<td style=\"text-align:left; vertical-align:top;\">";
                MessageHeader += "<img alt=\"NICE - National Institute for Health and Care Excellence\" hspace=0 src=\"cid:imageId\" align=baseline border=0>";
                MessageHeader += "</td>";
                MessageHeader += "<td style=\"text-align:right;\">";
                if (Header.ToString() == "Manchester")
                {
                    MessageHeader += ManchesterHeader;
                }
                else if (Header.ToString() == "London")
                {
                    MessageHeader += LondonHeader;
                }
                MessageHeader += "</td>";
                MessageHeader += "</tr>";
                MessageHeader += "</table>";
                MessageHeader += "</td>";
                MessageHeader += "</tr>";
            }
            MessageHeader += "<tr>";
            MessageHeader += "<td>";

            MessageFooter = "</td>";
            MessageFooter += "</tr>";
            MessageFooter += "</table>";
            MessageFooter += "</body>";
            MessageFooter += "</html>";

            //lblPreview.Text = MessageHeader + txtMessage.Text.ToString() + MessageFooter;
            //return;

            //about to start sending emails - update the MessageSendStatus to inprogress
            myCommand.CommandText = "P_EmailSendStatus_Update";
            myCommand.Parameters.Clear();
            myCommand.Parameters.Add("@P_EmailID", SqlDbType.Int).Value = P_EmailID;
            myCommand.Parameters.Add("@SendStatus", SqlDbType.TinyInt).Value = 1;
            myConnection.Open();
            myCommand.ExecuteNonQuery();
            myConnection.Close();

            bool ErrorHappened = false;
            LinkedResource imagelink;

            foreach (DataRow dr1 in dtRecipients.Rows)
            {
                Logger.Debug("Sending email to EmailRecipientID " + dr1["EmailRecipientID"].ToString());

                ErrorHappened = false;

                MailMessage message = new MailMessage();
                message.BodyEncoding = System.Text.Encoding.UTF8;
                EmailAddress = dr1["EmailAddress"].ToString();
                Recipient = dr1["Recipient"].ToString();

                string logopath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "images", "NICE-Master-72dpi-MIN.png");

                try
                {
                    imagelink = new LinkedResource(logopath, "image/png");
                }
                catch (Exception logoException)
                {
                    Logger.Error("Error loading logo image");
                    Logger.Error(logoException);
                    throw;
                }

                imagelink.ContentId = "imageId";
                imagelink.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;

                //create the views
                //AlternateView plainView = AlternateView.CreateAlternateViewFromString(Strip(HtmlBox1.Text.ToString()), null, "text/plain");
                AlternateView plainView = AlternateView.CreateAlternateViewFromString(MessagePlain.ToString().Replace("[Salutation]", Recipient), null, "text/plain");
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(MessageHeader + Message.ToString().Replace("[Salutation]", Recipient) + MessageFooter, null, "text/html");
                if (Header.ToString() == "Manchester" || Header.ToString() == "London")
                {
                    htmlView.LinkedResources.Add(imagelink);
                }
                Logger.Debug("Logo image added");

                //add the views
                message.AlternateViews.Add(plainView);
                message.AlternateViews.Add(htmlView);

                if (isEmail(EmailAddress) == true)
                {
                    Logger.Debug("Email address for EmailRecipientID " + dr1["EmailRecipientID"].ToString() + " is valid");
                    try
                    {
                        MailAddress fromAddress = new MailAddress(ReplyToAddress, "National Institute for Health and Care Excellence");
                        MailAddress toAddress = new MailAddress(EmailAddress, Recipient);
                        MailAddress replytoAddress = new MailAddress(ReplyToAddress, "National Institute for Health and Care Excellence");
                        message.From = fromAddress;
                        message.To.Add(toAddress);
                        message.IsBodyHtml = true;
                        message.ReplyToList.Add(replytoAddress);
                        message.Subject = Subject.ToString();

                        foreach (DataRow dr2 in dtAttachments.Rows)
                        {
                            if (File.Exists(dr2["Attachment"].ToString()))
                            {
                                message.Attachments.Add(new System.Net.Mail.Attachment(dr2["Attachment"].ToString()));
                            }
                        }

                        //throttle back
                        //Thread.Sleep(500);

                        smtpClient.Host = MailServer;
                        smtpClient.Send(message);
                        Logger.Debug("Email sent to EmailRecipientID " + dr1["EmailRecipientID"].ToString());
                    }
                    catch (SmtpException Smtpex)
                    {
                        ErrorHappened = true;
                        Logger.Error(Smtpex);
                        //write the exception to the database - already have connection and command objects
                        string ErrorMessage = Smtpex.Message.ToString();
                        if (ErrorMessage.Length >= 1000)
                        {
                            ErrorMessage = ErrorMessage.Substring(0, 999);
                        }

                        myCommand.CommandText = "P_EmailError_Add";
                        myCommand.Parameters.Clear();
                        myCommand.Parameters.Add("@P_EmailID", SqlDbType.Int).Value = P_EmailID;
                        myCommand.Parameters.Add("@EmailRecipientID", SqlDbType.Int).Value = Convert.ToInt32(dr1["EmailRecipientID"].ToString());
                        myCommand.Parameters.Add("@ErrorDescription", SqlDbType.VarChar, 1000).Value = ErrorMessage.ToString();
                        myConnection.Open();
                        myCommand.ExecuteNonQuery();
                        myConnection.Close();

                        continue;
                    }
                    catch (Exception ex)
                    {
                        ErrorHappened = true;
                        Logger.Error(ex);
                        //write the exception to the database - already have connection and command objects
                        string ErrorMessage = ex.Message.ToString();
                        if (ErrorMessage.Length >= 1000)
                        {
                            ErrorMessage = ErrorMessage.Substring(0, 999);
                        }

                        myCommand.CommandText = "P_EmailError_Add";
                        myCommand.Parameters.Clear();
                        myCommand.Parameters.Add("@P_EmailID", SqlDbType.Int).Value = P_EmailID;
                        myCommand.Parameters.Add("@EmailRecipientID", SqlDbType.Int).Value = Convert.ToInt32(dr1["EmailRecipientID"].ToString());
                        myCommand.Parameters.Add("@ErrorDescription", SqlDbType.VarChar, 1000).Value = ErrorMessage.ToString();
                        myConnection.Open();
                        myCommand.ExecuteNonQuery();
                        myConnection.Close();

                        continue;
                    }

                    finally
                    {
                        if (!ErrorHappened)
                        {
                            //to do - confirm by updating the Recipient Row using dr["DListEmailRecipientID"]
                            //already have connection and command objects
                            myCommand.CommandText = "P_EmailSentRecipient_Update";
                            myCommand.Parameters.Clear();
                            myCommand.Parameters.Add("@EmailRecipientID", SqlDbType.Int).Value = Convert.ToInt32(dr1["EmailRecipientID"].ToString());
                            myConnection.Open();
                            myCommand.ExecuteNonQuery();
                            myConnection.Close();
                        }
                    }
                }

                message.Dispose();
            }

            ds.Dispose();

            //finished sending emails - update the MessageSendStatus to finished
            myCommand.CommandText = "P_EmailSendStatus_Update";
            myCommand.Parameters.Clear();
            myCommand.Parameters.Add("@P_EmailID", SqlDbType.Int).Value = P_EmailID;
            myCommand.Parameters.Add("@SendStatus", SqlDbType.TinyInt).Value = 2;
            myConnection.Open();
            myCommand.ExecuteNonQuery();
            myConnection.Close();

            //did we attempt to contact any email in CNoEmail
            if (NumNoSend > 0)
            {
                MailMessage message = new MailMessage();
                message.BodyEncoding = System.Text.Encoding.UTF8;

                StringBuilder sb = new StringBuilder();

                sb.Append("<p>" + todaysDate + "</p>");
                sb.Append("<p>An attempt was made to send an email to an address on the Do Not Contact list.</p>");
                sb.Append("<p>This can be reviewed in the Contacts system via Manage | Do not Contact Attempts</p>");

                message.Body = sb.ToString();

                MailAddress fromAddress = new MailAddress("DoNotReply@nice.org.uk", "National Institute for Health and Care Excellence");
                message.From = fromAddress;
                MailAddress toAddress = new MailAddress(ReplyToAddress, User);
                message.To.Add(toAddress);
                MailAddress replytoAddress = new MailAddress("DoNotReply@nice.org.uk", "National Institute for Health and Care Excellence");
                message.IsBodyHtml = true;
                message.Subject = "Attempt to send email to Do Not Contact";

                smtpClient.Host = MailServer;
                smtpClient.Send(message);
            }

            smtpClient.Dispose();
            Logger.Debug("Finished calling P_WCF for environment " + whichapp + " and email " + P_EmailID.ToString());
        }

        private bool isEmail(string inputEmail)
        {
            if (!(String.IsNullOrEmpty(inputEmail)))
            {
                string strRegex = @"^([a-zA-Z0-9_\-\.\'\&]+)@((\[[0-9]{1,3}" +
                  @"\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" +
                  @".)+))([a-zA-Z0-9]*)(\]?)$";
                Regex re = new Regex(strRegex);
                if (re.IsMatch(inputEmail))
                    return (true);
                else
                    return (false);
            }
            else
            {
                return (false);
            }
        }
    }
}
