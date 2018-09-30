using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Exchange.WebServices.Data;


namespace ExchangeMailbox
{
    class Program
    {
        public static string mailuser = string.Empty;
        public static string mailpwd = string.Empty;
        public static string maildomain = string.Empty;
        public static string mailservice = string.Empty;
        public static string mailserver = string.Empty;
        public static string LogFile;

        //This function will load all the Exchange configuration/credentials from Settings.xml
        public static void LoadSettings()
        {
            string fileName = "Settings.xml";
            
            if(File.Exists(fileName)) 
            {
                XDocument xmlDoc = XDocument.Load(fileName);

                mailuser = xmlDoc.Root.Element("mail_user").Value;
                mailpwd = xmlDoc.Root.Element("mail_pwd").Value;
                maildomain = xmlDoc.Root.Element("mail_domain").Value;
                mailservice = xmlDoc.Root.Element("mail_service").Value;
                mailserver = xmlDoc.Root.Element("mail_server").Value;
                LogFile = xmlDoc.Root.Element("LogFile").Value;

                WriteLog("Loading Process Settings");
                WriteLog("Process Settings loaded successfully");
            }
            else
            {
                WriteLog("Process Settings not found. Aborting process");
                Environment.Exit(0);
            }
        }

        public static void WriteLog(string log) 
        {
            Console.WriteLine(DateTime.Now.ToString() + ":      " + log);
            File.AppendAllText(LogFile, DateTime.Now.ToString() + ":    " + log + Environment.NewLine);
        }

        public void Start()
        {
            LoadSettings();
            WriteLog("Job Started");
            WriteLog("Creating EWS Connection");
            ExchangeService service = new ExchangeService();
            service.Timeout = 600000;
            service.Credentials = new NetworkCredential(mailuser, mailpwd, maildomain);
            service.Url = new Uri(mailservice);
            WriteLog("EWS Connection Created");
            
            WriteLog("Creating Event Subscription");
            WriteLog("Event Subscription Started");
            
            ItemView view = new ItemView(10);
            FindItemsResults<Item> findResults;

            WriteLog("Checking Existing Emails");
            do
            {
                //To check count of emails in Inbox folder.
                //findResults = service.FindItems(WellKnownFolderName.Inbox, view);

                //To check count of emails in Junk folder.
                findResults = service.FindItems(WellKnownFolderName.JunkEmail, view);

                int Total = findResults.Count();
                WriteLog("There are " + Total + " Emails in Junk Folder");
                WriteLog("========== Run Completed ==========");

            } while (findResults.MoreAvailable);

        }

        static void Main(string[] args)
        {
            Program phase1 = new Program();
            phase1.Start();
            Console.ReadLine();
        }
    }
}
