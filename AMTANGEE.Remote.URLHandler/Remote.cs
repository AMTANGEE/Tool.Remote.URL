using System;
using System.Collections.Generic;
using System.Text;

namespace AMTANGEE
{
    public partial class Remote
    {
        public class Tools
        {

            public static AMTANGEE.SDK.Contacts.ContactBase GetContactBySource(string source, string adressType)
            {
                AMTANGEE.SDK.Contacts.ContactBase cb = null;
                if (adressType.ToLower() == "person")
                {

                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchBySource(AMTANGEE.SDK.Contacts.SearchKinds.EndsWith, ":" + source);
                    if (contacts.Count > 0)
                        cb = contacts[0];
                }
                else
                {
                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchBySource(AMTANGEE.SDK.Contacts.SearchKinds.EndsWith, ":" + source);
                    if (contacts.Count > 0)
                    {
                        foreach (AMTANGEE.SDK.Contacts.ContactBase contactBase in contacts)
                        {
                            if (contactBase is AMTANGEE.SDK.Contacts.Contact)
                            {
                                cb = contactBase;
                                break;
                            }
                        }
                    }
                }
                return cb;
            }

            public static AMTANGEE.SDK.Contacts.ContactBase GetContactByCustomerOrVendorNo(string customerOrVendorNo,string adressType)
            {
                AMTANGEE.SDK.Contacts.ContactBase cb = null;
                AMTANGEE.SDK.Contacts.Search s = new SDK.Contacts.Search(AMTANGEE.SDK.Global.CurrentUser);
                s.SearchCriteria = "ADDRESSES.DLL:KUNDENNUMMER=" + customerOrVendorNo;
                if (adressType.ToLower() == "person")
                {

                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.Search(s, false, false); //AMTANGEE.SDK.Contacts.Contacts.SearchBySource(AMTANGEE.SDK.Contacts.SearchTypes.EndsWith, ":" + parameter["number"]);
                    if (contacts.Count > 0)
                        cb = contacts[0];
                }
                else
                {
                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.Search(s, false, false);//AMTANGEE.SDK.Contacts.Contacts.SearchBySource(AMTANGEE.SDK.Contacts.SearchTypes.EndsWith, ":" + parameter["number"]);
                    if (contacts.Count > 0)
                    {
                        foreach (AMTANGEE.SDK.Contacts.ContactBase contactBase in contacts)
                        {
                            if (contactBase is AMTANGEE.SDK.Contacts.Contact)
                            {
                                cb = contactBase;
                            }
                        }
                    }
                }

                return cb;
            }

            public static string IfEmptyString(string input)
            {
                return IfEmptyString(input, "");
            }

            public static string IfEmptyString(string input,string defaultValue)
            {
                if (input != null)
                    return input;

                return defaultValue;
            }
        
            public static AMTANGEE.SDK.Settings.Helper Settings
            {
                get
                {
                    return AMTANGEE.SDK.Settings.PerUser["AMTANGEEURL"];
                }

            }

            private static string GetCommandLine(System.Diagnostics.Process process)
            {
                var commandLine = new StringBuilder(process.MainModule.FileName + " ");

                using (var searcher = new System.Management.ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + process.Id))
                {
                    foreach (var @object in searcher.Get())
                    {
                        commandLine.Append(@object["CommandLine"] + " ");
                    }
                }
                return commandLine.ToString();
            }

            public static bool CheckForAMTANGEE(bool force)
            {
                if (force)
                    return true;

                return CheckForAMTANGEE();
            }

            public static bool CheckForAMTANGEE()
            {
                bool found = false;

                System.Diagnostics.Process[] runningProcesses = System.Diagnostics.Process.GetProcesses();
                var currentSessionID = System.Diagnostics.Process.GetCurrentProcess().SessionId;

                foreach (var p in runningProcesses)
                {
                    if (p.SessionId == currentSessionID)
                    {
                        if (p.ProcessName.ToUpper() == "AMTANGEE")
                        {
                            found = true;
                            break;
                        }
                    }
                }

                if (!found)
                {
                    foreach (System.Diagnostics.Process process in runningProcesses)
                    {
                        if (process.SessionId == currentSessionID)
                        {
                            if (process.ProcessName.ToUpper().StartsWith("AMTANGEE ") || process.MainWindowTitle.ToUpper().StartsWith("AMTANGEE® -") || process.MainWindowTitle.ToUpper().StartsWith("AMTANGEE® MOBILE") || process.MainWindowTitle.ToUpper().StartsWith("AMTANGEE® BRANCH") || process.MainWindowTitle.ToUpper().StartsWith("AMTANGEE® SMALL BRANCH"))
                            {
                                found = true;
                                break;
                            }
                        }
                    }
                }

                if (!found)
                {
                    System.Windows.Forms.MessageBox.Show("Um diese Funktion nutzen zu können muss AMTANGEE® gestartet sein!", "Hinweis", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }
                return true;
            }

            public static AMTANGEE.SDK.ICategoryObject GetCategoryByFullName(AMTANGEE.SDK.ICategoryObject parent, string name)
            {
                AMTANGEE.SDK.ICategoryObject result = null;

                if (parent.FullName.ToUpper() == name.ToUpper())
                    return parent;

                foreach (AMTANGEE.SDK.ICategoryObject cat in parent.Children)
                {
                    result = GetCategoryByFullName(cat, name);
                }

                return result;
            }
        }

        public class RemoteUrl
        {
            public static void Exec(string module, System.Collections.Specialized.NameValueCollection urlParameters)
            {
                Paramters parameters = new Paramters();
                foreach (string str in urlParameters.AllKeys)
                {
                    parameters.AddValue(str, urlParameters[str]);
                }

                switch (module.ToUpper())
                {
                    case "SMS":
                        AMTANGEE.Remote.DefaultHandler.SMS sms = new AMTANGEE.Remote.DefaultHandler.SMS();
                        sms.Exec(parameters);
                        break;
                    case "CALL":
                        AMTANGEE.Remote.DefaultHandler.Call call = new AMTANGEE.Remote.DefaultHandler.Call();
                        call.Exec(parameters);
                        break;
                    case "PROJECT":
                        AMTANGEE.Remote.DefaultHandler.Project project = new AMTANGEE.Remote.DefaultHandler.Project();
                        project.Exec(parameters);
                        break;
                    case "GENERAL":
                        AMTANGEE.Remote.DefaultHandler.General general = new AMTANGEE.Remote.DefaultHandler.General();
                        general.Exec(parameters);
                        break;
                    case "DOCUMENT":
                        AMTANGEE.Remote.DefaultHandler.Document document = new AMTANGEE.Remote.DefaultHandler.Document();
                        document.Exec(parameters);
                        break;
                    case "APPOINTMENT":
                        AMTANGEE.Remote.DefaultHandler.Appointment appointment = new Remote.DefaultHandler.Appointment();
                        appointment.Exec(parameters);
                        break;
                    case "CONTACT":
                        AMTANGEE.Remote.DefaultHandler.Contact contact = new Remote.DefaultHandler.Contact();
                        contact.Exec(parameters);
                        break;
                    case "EMAIL":
                        AMTANGEE.Remote.DefaultHandler.Email email = new Remote.DefaultHandler.Email();
                        email.Exec(parameters);
                        break;
                }
            }
        }

        public class Paramters : System.Collections.Generic.Dictionary<string, string>
        {
            public void AddValue(string name, string value)
            {
                name = name.ToUpper();
                base.Add(name, value);
            }

            public string GetValue(string key)
            {
                key = key.ToUpper();
                if (this.ContainsKey(key))
                    return base[key];

                return "";
            }

            public string this[string key]
            {
                get
                {
                    key = key.ToUpper();
                    if (this.ContainsKey(key))
                        return base[key];
                    else
                        return "";
                }
                set
                {
                    key = key.ToUpper();
                    base[key] = value;
                }
            }
        }

        public interface IRemoteHandler
        {
            void Exec(Paramters parameters);

            string ModuleName
            {
                get;
            }
        }
    }
}
