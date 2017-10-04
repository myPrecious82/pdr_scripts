using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace pdr_scripts
{
    public class Program
    {
        public static void Main()
        {
            const string path = "\\\\cfco06nds01\\isd01\\SysDev\\CANTS\\PDR RUN SQL";
            const string createdBy = "DCFS\\CAA6377";
            const string emailSmtpHost = "SMTPR.illinois.gov";
            string emailSendTo = "Cuneyt.Barutcu@illinois.gov";
#if DEBUG
            emailSendTo = "alexis.atchison@illinois.gov";
#endif
            const string emailSendFrom = "alexis.atchison@illinois.gov";
            const string emailSendFromDisplay = "Atchison, Alexis";
            const string crlf = "<br />";

            var scripts = Directory.EnumerateFiles(path, "vantive*.sql", SearchOption.TopDirectoryOnly).ToList();
            var cutoffDate = DateTime.Now.Date;
#if DEBUG
            cutoffDate = DateTime.Now.Date.AddDays(-3);
#endif
            var emailSubject = $"{cutoffDate.ToString("MM.dd.yyyy")} PDR Scripts";

            var intEvlStDate = $"Intake Evaluations - change start date{crlf}";
            var invstFlipScr = $"Investigations - flip SCR{crlf}";
            var invstMakeFac = $"Investigations - make facility{crlf}";
            var invstEndDupe = $"Investigations - end duplicate assignment{crlf}";
            var invstChgStrt = $"Investigations - change assignment start date{crlf}";
            var personMerge = $"Person Merge{crlf}";
            var other = $"Other{crlf}";

            var fileCount = 0;

            foreach (var script in scripts)
            {
                var myFileCreateTs = File.GetCreationTime(script);
                var myFileModifyTs = File.GetLastWriteTime(script);
                if (myFileCreateTs <= cutoffDate && myFileModifyTs <= cutoffDate) continue;

                var attr = File.GetAccessControl(script).GetOwner(typeof(System.Security.Principal.NTAccount)).ToString().ToUpper();
                if (!attr.Equals(createdBy)) continue;

                // this is a desired file, so strip down to just the ticket number and read line 3 to see what type of ticket it is
                var vantiveNum = script.ToLower().Remove(0, path.Length + 1).Replace("vantive ", "").Replace(".sql", "");
                var line3 = File.ReadLines(script).Skip(2).Take(1).First();
                fileCount++;
                var url = $"<a href='http://helpdesk/Users/Ticket.aspx?TicketId={vantiveNum.Substring(0, 7)}'>{vantiveNum}</a>";

                switch (line3.ToLower())
                {
                    case "--\tintake evaluations - change start date":
                        intEvlStDate = $"{intEvlStDate}{url}{crlf}";
                        break;
                    case "--\tinvestigations - flip scr":
                        invstFlipScr  = $"{invstFlipScr}{url}{crlf}";
                        break;
                    case "--\tinvestigations - make facility":
                        invstMakeFac  = $"{invstMakeFac}{url}{crlf}";
                        break;
                    case "--\tinvestigations - end duplicate assignment":
                        invstEndDupe  = $"{invstEndDupe}{url}{crlf}";
                        break;
                    case "--\tinvestigations - change assignment start date":
                        invstChgStrt  = $"{invstChgStrt}{url}{crlf}";
                        break;
                    case "--\tperson merge":
                        personMerge  = $"{personMerge}{url}{crlf}";
                        break;
                    default:
                        other  = $"{other}{url}{crlf}";
                        break;
                }
            }
            // finished looping through and gathering pertinent files
            if (fileCount <= 0)
            {
                return;
            }

            intEvlStDate = (intEvlStDate.ToLower().Equals($"intake evaluations - change start date{crlf}"))
                ? string.Empty
                : string.Concat(intEvlStDate, crlf);

            invstFlipScr = (invstFlipScr.ToLower().Equals($"investigations - flip scr{crlf}"))
                ? string.Empty
                : string.Concat(invstFlipScr, crlf);

            invstMakeFac = (invstMakeFac.ToLower().Equals($"investigations - make facility{crlf}"))
                ? string.Empty
                : string.Concat(invstMakeFac, crlf);

            invstEndDupe = (invstEndDupe.ToLower().Equals($"investigations - end duplicate assignment{crlf}"))
                ? string.Empty
                : string.Concat(invstEndDupe, crlf);

            invstChgStrt = (invstChgStrt.ToLower().Equals($"investigations - change assignment start date{crlf}"))
                ? string.Empty
                : string.Concat(invstChgStrt, crlf);

            personMerge = (personMerge.ToLower().Equals($"person merge{crlf}"))
                ? string.Empty
                : string.Concat(personMerge, crlf);

            other = (other.Equals($"Other{crlf}"))
                ? string.Empty
                : string.Concat(other, crlf);

            var client = new SmtpClient(emailSmtpHost);

            var from = new MailAddress(emailSendFrom, emailSendFromDisplay, System.Text.Encoding.UTF8);
            var to = new MailAddress(emailSendTo);
            var message = new MailMessage(from, to)
            {
                IsBodyHtml = true,
                BodyEncoding = System.Text.Encoding.UTF8,
                Subject = emailSubject,
                SubjectEncoding = System.Text.Encoding.UTF8,
                Bcc = { new MailAddress(emailSendFrom) }
            };
            message.Body = $"{message.Body}<span style='font-size:11pt;font-family:Calibri'>{intEvlStDate}{invstFlipScr}{invstMakeFac}{invstEndDupe}{invstChgStrt}{personMerge}{other}</span>";
            
            client.Send(message);

            message.Dispose();
        }
    }
}

