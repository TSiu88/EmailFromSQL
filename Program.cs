using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using MimeKit;
using MailKit;
using MailKit.Net.Smtp;
using System.Data.SqlClient;

namespace EmailFromSQL
{
    class Program
    {
        protected static string HostIP = ConfigurationManager.AppSettings["SMTPip"].ToString();
        protected static string Username = ConfigurationManager.AppSettings["SMTPusername"].ToString();
        protected static string Pass = ConfigurationManager.AppSettings["SMTPpassword"].ToString();
        protected static int HostPort = Int32.Parse(ConfigurationManager.AppSettings["SMTPport"]);
        protected static string FromEmail = ConfigurationManager.AppSettings["EmailFrom"].ToString();
        protected static string Recipient = ConfigurationManager.AppSettings["EmailNotificationTo"].ToString();
        protected static string SubjectStr = "Email from SQL Test";

        private static readonly string CRCCon = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
        private static string query = "SELECT TOP(10) ItemNumber , I.ItemDescription, MOQ, MPQ, CreateDate, LastUpdate FROM PIM_Replen_Attr RA join BP_ITEM_MASTER I ON RA.ItemNumber = I.itemlookupcode";

        static void Main()
        {
            MimeMessage mimeMail = CreateNewMimeMessage();
            SendMimeMessage(mimeMail);

            Console.WriteLine("Mail Sent!");
            Console.ReadLine();
        }

        private static MimeMessage CreateNewMimeMessage()
        {
            MimeMessage mm = new MimeMessage();
            mm.From.Add(MailboxAddress.Parse(FromEmail));
            mm.Subject = SubjectStr;
            mm.To.Add(MailboxAddress.Parse(Recipient));
            BodyBuilder bb = new BodyBuilder();
            bb.HtmlBody = GetMailMessage();
            mm.Body = bb.ToMessageBody();
            return mm;
        }

        private static string GetMailMessage()
        {
            string SqlString = GetSqlHTML();
            StringBuilder HTMLBody = new StringBuilder();
            HTMLBody.Append("<HTML>");
            HTMLBody.Append("<HEAD>");
            HTMLBody.Append("<TITLE>" + SubjectStr +"</TITLE>");
            HTMLBody.Append("</HEAD>");
            HTMLBody.Append("<BODY>");
            HTMLBody.Append("<H1>First 10 IN PIM_Replen_Attr Table</H1>");
            HTMLBody.Append("<H2>Testing Removing each row from SQL server</H2>");
            
            HTMLBody.Append("<TABLE width='1000px' border='1' cellpadding=5 cellspacing=0>");
            HTMLBody.Append("<TR>");

            HTMLBody.Append("<TD width='120px'>");
            HTMLBody.Append("SKU");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='380px'>");
            HTMLBody.Append("Item Description");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='50px'>");
            HTMLBody.Append("MOQ");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='50px'>");
            HTMLBody.Append("MPQ");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='200px'>");
            HTMLBody.Append("Create Date");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("<TD width='200px'>");
            HTMLBody.Append("Last Update");
            HTMLBody.Append("</TD>");

            HTMLBody.Append("</TR>");

            HTMLBody.Append(SqlString);

            HTMLBody.Append("</TABLE>");
            HTMLBody.Append("</BODY>");
            HTMLBody.Append("</HTML>");
            return HTMLBody.ToString();
        }

        public static string GetSqlHTML()
        {
            string output = "";

            using (SqlConnection db = new SqlConnection(CRCCon))
            {
                try
                {
                    db.Open();
                    Console.WriteLine("Connection to server successful!");
                    try
                    {
                        SqlCommand command = new SqlCommand(query, db);
                        SqlDataReader reader = command.ExecuteReader();

                        if(reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                Object[] values = new Object[reader.FieldCount];
                                int count = reader.GetValues(values);
                                output += "<tr>";
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    output += "<td>" + values[i].ToString().Trim() + "</td>";
                                }
                                output += "</tr>";
                            }
                            
                        }
                    }
                    catch (SqlException ex)
                    {
                        Console.WriteLine("Query failed: " + ex.Message);
                    }
                }
                catch (SqlException ex)
                {
                    Console.WriteLine("Connection failed: " + ex.Message);
                }
                finally
                {
                    Console.WriteLine(output);
                    db.Close();
                    db.Dispose();
                }
            }

            return output;
        }

        private static void SendMimeMessage (MimeMessage mail)
        {
            var client = new MailKit.Net.Smtp.SmtpClient();
            client.Connect(HostIP, HostPort, false);
            client.Authenticate(Username, Pass);
            client.Send(mail);
            client.Disconnect(true);
        }
    }
}
