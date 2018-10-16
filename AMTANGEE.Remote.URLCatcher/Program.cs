using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualBasic.ApplicationServices;
using IWshRuntimeLibrary;
using Microsoft.Win32;
using System.Security.Permissions;
using System.Security.Principal;
namespace AMTANGEEUrl
{
    static class Program
    {

        public class SingleInstanceController : WindowsFormsApplicationBase
        {
            public SingleInstanceController()
            {
                IsSingleInstance = true;
                StartupNextInstance += this_StartupNextInstance;
            }


            void this_StartupNextInstance(object sender, StartupNextInstanceEventArgs e)
            {
                frmMain form = MainForm as frmMain;
                MainForm.ShowInTaskbar = false;
                MainForm.WindowState = FormWindowState.Minimized;
                MainForm.Visible = false;
                if (e.CommandLine.Count > 0)
                {
                    _param = e.CommandLine[0];
                    _param2 = new string[e.CommandLine.Count];
                    int i = 0;
                    foreach (string str in e.CommandLine)
                    {
                        _param2[i] = str;
                        i++;
                    }
                    if (_param2.Length > 0 && _param2[0].ToUpper() == "CLOSE")
                    {
                        System.Diagnostics.Process.GetCurrentProcess().Kill();
                    }
                }
                else
                {
                    _param = "";
                    _param2 = new string[0];
                }
                Do();
            }

            protected override void OnCreateMainForm()
            {
                MainForm = new frmMain();
                MainForm.ShowInTaskbar = false;
                MainForm.WindowState = FormWindowState.Minimized;
                MainForm.Visible = false;
                Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\AMTANGEE\CRM\Current");
                string userID = "";
                if (regKey != null)
                {
                    object UserGuid = regKey.GetValue("User");
                    if (UserGuid != null)
                    {
                        userID = UserGuid.ToString().Replace("{", "").Replace("}", "");
                    }
                }
                int dbNumber = 0;
                regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\AMTANGEE\RemoteUrl");
                if (regKey != null)
                {
                    string dbNumberString = regKey.GetValue("DefaultDB","0").ToString();
                    try
                    {
                        dbNumber = Convert.ToInt16(dbNumberString);
                    }
                    catch
                    {
                    }
                }
                if (dbNumber > 0)
                    AMTANGEE.DB.Open(dbNumber);
                else
                    AMTANGEE.DB.Open();
                
                if (userID.Length == 0)
                {
                    AMTANGEE.SDK.Users.Users.Login();
                    userID = AMTANGEE.SDK.Users.Users.Current.Guid.ToString();
                }
                else
                {
                    try
                    {
                        AMTANGEE.SDK.Users.Users.Login(new AMTANGEE.SDK.Users.User(new Guid(userID)));
                    }
                    catch
                    {
                        AMTANGEE.SDK.Users.Users.Login();
                        userID = AMTANGEE.SDK.Users.Users.Current.Guid.ToString();
                    }
                }
                _userID = userID;
                Do();
            }
        }


        static void ImportDocuments(System.IO.DirectoryInfo directory, AMTANGEE.SDK.Documents.Category parentCat, AMTANGEE.Forms.ProgressForm pf)
        {
            AMTANGEE.SDK.Documents.Category docCat = null;
            foreach (AMTANGEE.SDK.Documents.Category c in parentCat.Children)
            {
                if (c.Name.ToUpper() == directory.Name.ToUpper())
                {
                    docCat = c;
                    break;
                }
            }
            if (docCat == null)
            {
                docCat = new AMTANGEE.SDK.Documents.Category(parentCat);
                docCat.Name = directory.Name;
                docCat.Save();
            }
            System.IO.FileInfo[] files = directory.GetFiles();
            foreach (System.IO.FileInfo file in files)
            {
                pf.Caption = file.FullName;
                AMTANGEE.SDK.Documents.Document d = new AMTANGEE.SDK.Documents.Document(docCat);
                d.DateTime = DateTime.Now;
                d.Description = file.Name;
                if (file.Extension.Trim().Length <= 5)
                    d.FileType = file.Extension;
                else
                    d.FileType = file.Extension.Substring(0, 5);
                d.LoadFromFile(file.FullName);
                d.Save();
                pf.Position++;
                Application.DoEvents();
            }
            System.IO.DirectoryInfo[] directories = directory.GetDirectories();
            foreach (System.IO.DirectoryInfo d in directories)
            {
                ImportDocuments(d, docCat,pf);
            }
        }


        
        static void TimerTick(object state)
        {
            vTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);
            System.Diagnostics.Process[] pname = System.Diagnostics.Process.GetProcessesByName("loader");
            if (pname.Length > 0)
            {
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }
            else
                vTimer.Change(1000,System.Threading.Timeout.Infinite);
        }


        static System.Threading.TimerCallback vCallback = null;
        static System.Threading.Timer vTimer = null;
       static bool isDirectory(string path)
        {
            System.IO.FileAttributes attr = System.IO.File.GetAttributes(path);

            if ((attr & System.IO.FileAttributes.Directory) == System.IO.FileAttributes.Directory)
                return true;

            return false;
                
        }
        static string _userID = "";
        static bool started = false;
        static string _param;
        static string _param_2;
        static string[] _param2;
        static void Do()
        {
            if (_param == null)
                _param = "";

            if (_param_2 == null)
                _param_2 = "";

            if (vTimer == null)
            {
                vCallback = new System.Threading.TimerCallback(TimerTick);
                vTimer = new System.Threading.Timer(vCallback, null, 1000, 1000);
            }
            if (_param.Length > 0)
            {
                  string parameter = _param;
                try
                {
                    if (!parameter.ToLower().StartsWith("amtangee://"))
                    {
                        if (_param2 != null && _param2.Length > 1 && parameter == "EML" && _param2[1].Trim().Length > 0)
                        {
                            if (!AMTANGEE.Remote.Tools.CheckForAMTANGEE())
                                return;
                            try
                            {
                                Guid guid = Guid.NewGuid();
                                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\AMTANGEE\\EMLViewer", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key.SetValue(guid.ToString() + "_File", _param2[1], Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue(guid.ToString() + "_User", System.Environment.UserName, Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue(guid.ToString() + "_Date", DateTime.Now.ToString(), Microsoft.Win32.RegistryValueKind.String);
                                AMTANGEE.SDK.Events.SendSDKEvent("SHOWEML:" + guid.ToString(),false);
                            }
                            catch
                            {
                            }
                        }

                        if (_param2 != null && _param2.Length > 1 && parameter == "OPENEML" && _param2[1].Trim().Length > 0)
                        {
                            if (!AMTANGEE.Remote.Tools.CheckForAMTANGEE())
                                return;
                            try
                            {
                                Guid guid = Guid.NewGuid();
                                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\AMTANGEE\\FileViewer", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key = key.CreateSubKey( guid.ToString(), RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key.SetValue( "File", _param2[1], Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue( "User", System.Environment.UserName, Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue( "Date", DateTime.Now.ToString(), Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Type", "EML", Microsoft.Win32.RegistryValueKind.String);
                                AMTANGEE.SDK.Events.Send(AMTANGEE.SDK.Global.CurrentUser, "*AMTANGEE.MODULES.MESSAGES.DLL:OPENEML:{" + guid.ToString().ToUpper() + "}", true);
                            }
                            catch
                            {
                            }
                        }

                        if (_param2 != null && _param2.Length > 1 && parameter == "OPENMSG" && _param2[1].Trim().Length > 0)
                        {
                            if (!AMTANGEE.Remote.Tools.CheckForAMTANGEE())
                                return;
                            try
                            {
                                Guid guid = Guid.NewGuid();
                                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\AMTANGEE\\FileViewer", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key = key.CreateSubKey(guid.ToString(), RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key.SetValue("File", _param2[1], Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("User", System.Environment.UserName, Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Date", DateTime.Now.ToString(), Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Type", "MSG", Microsoft.Win32.RegistryValueKind.String);
                                AMTANGEE.SDK.Events.Send(AMTANGEE.SDK.Global.CurrentUser, "*AMTANGEE.MODULES.MESSAGES.DLL:OPENMSG:{" + guid.ToString().ToUpper() + "}", true);
                            }
                            catch
                            {
                            }
                        }

                        if (_param2 != null && _param2.Length > 1 && parameter == "OPENVCS" && _param2[1].Trim().Length > 0)
                        {
                            if (!AMTANGEE.Remote.Tools.CheckForAMTANGEE())
                                return;
                            try
                            {
                                Guid guid = Guid.NewGuid();
                                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\AMTANGEE\\FileViewer", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key = key.CreateSubKey(guid.ToString(), RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key.SetValue("File", _param2[1], Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("User", System.Environment.UserName, Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Date", DateTime.Now.ToString(), Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Type", "VCS", Microsoft.Win32.RegistryValueKind.String);
                                AMTANGEE.SDK.Events.Send(AMTANGEE.SDK.Global.CurrentUser, "*AMTANGEE.MODULES.APPOINTMENTS.DLL:OPENVCS:{" + guid.ToString().ToUpper() + "}", true);
                            }
                            catch
                            {
                            }
                        }

                        if (_param2 != null && _param2.Length > 1 && parameter == "OPENICS" && _param2[1].Trim().Length > 0)
                        {
                            if (!AMTANGEE.Remote.Tools.CheckForAMTANGEE())
                                return;
                            try
                            {
                                Guid guid = Guid.NewGuid();
                                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\AMTANGEE\\FileViewer", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key = key.CreateSubKey(guid.ToString(), RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key.SetValue("File", _param2[1], Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("User", System.Environment.UserName, Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Date", DateTime.Now.ToString(), Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Type", "ICS", Microsoft.Win32.RegistryValueKind.String);
                                AMTANGEE.SDK.Events.Send(AMTANGEE.SDK.Global.CurrentUser, "*AMTANGEE.MODULES.APPOINTMENTS.DLL:OPENICS:{" + guid.ToString().ToUpper() + "}", true);
                            }
                            catch
                            {
                            }
                        }

                        if (_param2 != null && _param2.Length > 1 && parameter == "OPENVCF" && _param2[1].Trim().Length > 0)
                        {
                            if (!AMTANGEE.Remote.Tools.CheckForAMTANGEE())
                                return;
                            try
                            {
                                Guid guid = Guid.NewGuid();
                                Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\\AMTANGEE\\FileViewer", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key = key.CreateSubKey(guid.ToString(), RegistryKeyPermissionCheck.ReadWriteSubTree);
                                key.SetValue("File", _param2[1], Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("User", System.Environment.UserName, Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Date", DateTime.Now.ToString(), Microsoft.Win32.RegistryValueKind.String);
                                key.SetValue("Type", "VCF", Microsoft.Win32.RegistryValueKind.String);
                                AMTANGEE.SDK.Events.Send(AMTANGEE.SDK.Global.CurrentUser, "*AMTANGEE.MODULES.CONTACTS.DLL:OPENVCF:{" + guid.ToString().ToUpper() + "}", true);
                            }
                            catch
                            {
                            }
                        }


                        if (_param2 != null && _param2.Length == 1)
                        {
                            try
                            {
                                if (System.IO.File.Exists(_param2[0]))
                                {
                                    System.IO.FileInfo fi = new System.IO.FileInfo(_param2[0]);


                                    bool found = false;
                                    switch (fi.Extension.Replace(".", "").ToUpper())
                                    {
                                        case "EML":
                                            SetAssociation(".eml", "EML_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENEML", "4", "EML-Datei");
                                            found = true;
                                            break;

                                        case "MSG":
                                            SetAssociation(".msg", "MSG_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENMSG", "4", "MSG-Datei");
                                            found = true;
                                            break;

                                        case "VCF":
                                            SetAssociation(".vcf", "VCF_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENVCF", "5", "VCF-Datei");
                                            found = true;
                                            break;

                                        case "ICS":
                                            SetAssociation(".ics", "ICS_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENICS", "6", "ICS-Datei");
                                            found = true;
                                            break;

                                        case "VCS":
                                            SetAssociation(".vcs", "VCS_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENVCS", "6", "VCS-Datei");
                                            found = true;
                                            break;
                                    }

                                    if (found)
                                        System.Diagnostics.Process.Start(_param2[0]);
                                }
                            }
                            catch
                            {
                            }
                        }

                        try
                        {
                            AMTANGEE.Forms.SelectCategoryForm scf = new AMTANGEE.Forms.SelectCategoryForm(AMTANGEE.Controls.CategoriesTree.Kinds.Documents, AMTANGEE.SDK.CategoryRights.Add, "Zielkategorie", "Ok");
                            try
                            {
                                foreach (string str in _param2)
                                {
                                    if (System.IO.Directory.GetDirectories(str).Length > 0)
                                    {
                                        AMTANGEE.SDK.Documents.Categories cats = new AMTANGEE.SDK.Documents.Categories();
                                        foreach (AMTANGEE.SDK.Documents.Category c in cats)
                                        {
                                            if (!c.HasRight(AMTANGEE.SDK.CategoryRights.CreateSubCategory))
                                            {
                                                scf.BlockCategory(c);
                                            }
                                        }
                                    }
                                }

                                if (scf.ShowDialog() == DialogResult.OK)
                                {
                                    AMTANGEE.SDK.Documents.Category docCat = (AMTANGEE.SDK.Documents.Category)scf.SelectedCategory;
                                    Environment.GetFolderPath(Environment.SpecialFolder.SendTo);
                                    AMTANGEE.Forms.ProgressForm pf = new AMTANGEE.Forms.ProgressForm();
                                    pf.Text = "Dateien werden importiert...";
                                    pf.StartPosition = FormStartPosition.CenterScreen;
                                    pf.TopMost = true;
                                    pf.Caption = "";
                                    int allFiles = 0;
                                    System.IO.DirectoryInfo di = null;
                                    foreach (string p in _param2)
                                    {
                                        if (isDirectory(p))
                                        {
                                            allFiles += new System.IO.DirectoryInfo(p).GetFiles("*.*", System.IO.SearchOption.AllDirectories).Length;
                                        }
                                        else
                                        {
                                            allFiles++;
                                        }
                                    }

                                    pf.Maximum = allFiles;
                                    pf.Show();
                                    AMTANGEE.SDK.Documents.Documents.BeginUpdate();
                                    foreach (string p in _param2)
                                    {
                                        di = null;
                                        System.IO.FileInfo fi = null;
                                        if (isDirectory(p))
                                        {
                                            di = new System.IO.DirectoryInfo(p);
                                        }
                                        else
                                        {
                                            fi = new System.IO.FileInfo(p);
                                        }

                                        if (di != null)
                                        {
                                            ImportDocuments(di, docCat, pf);
                                        }

                                        if (fi != null)
                                        {
                                            pf.Caption = fi.FullName;
                                            AMTANGEE.SDK.Documents.Document d = new AMTANGEE.SDK.Documents.Document(docCat);
                                            d.DateTime = DateTime.Now;
                                            d.Description = fi.Name;
                                            if (fi.Extension.Trim().Length <= 5)
                                                d.FileType = fi.Extension;
                                            else
                                                d.FileType = fi.Extension.Substring(0, 5);
                                            d.LoadFromFile(fi.FullName);
                                            d.Save();
                                            pf.Position++;
                                            Application.DoEvents();
                                        }
                                    }
                                    pf.Hide();
                                }
                            }
                            catch
                            {
                            }
                        }
                        catch (Exception exc)
                        {
                            string test = exc.Message;
                        }
                        finally
                        {
                            AMTANGEE.SDK.Documents.Documents.EndUpdate();
                        }
                    }
                }
                catch(Exception exc)
                {
                }

                if (parameter.ToLower().StartsWith("amtangee://"))
                    parameter = parameter.Substring(11);

                AMTANGEE.SDK.Settings.Helper settingsObject = AMTANGEE.SDK.Settings.PerUser["AMTANGEEURL"];

                string module = AMTANGEE.SDK.Global.CopyToPattern(parameter, '/');
                parameter = AMTANGEE.SDK.Global.CopyFromPattern(parameter, '/');
                System.Collections.Specialized.NameValueCollection nvc = System.Web.HttpUtility.ParseQueryString(parameter);
                if (module == "general")
                {
                    if (GetOption(parameter, "action") == "close")
                    {
                        System.Diagnostics.Process.GetCurrentProcess().Kill();
                        return;
                    }
                }

                nvc.Add("[USERGUID]", _userID);
                string referenzGUID = "";
                string referenzID = "";
                if (GetOption(parameter, "reference").Trim().Length > 0)
                {
                    referenzID = GetOption(parameter, "reference").ToUpper().Trim();
                    referenzGUID = settingsObject[GetOption(parameter, "reference").ToUpper().Trim()];
                }

                nvc.Add("[REF]", referenzID);
                nvc.Add("[REFGUID]", referenzGUID);

                if (module.ToLower() == "version")
                {
                    MessageBox.Show("AMTANGEE® Remote Url Version 3.0");
                    return;
                }

                if (module.ToLower() == "licenseholder")
                {
                    MessageBox.Show("Lizensiert für: "  + AMTANGEE.SDK.License.LicensedTo);
                    return;
                }

                AMTANGEE.Remote.RemoteUrl.Exec(module, nvc);
            }
        }

        static DateTime lastOpen = DateTime.MinValue;
        static void RegisterUrl()
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey("amtangee", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);
            key.SetValue("", "URL:amtangee (amtangee protocol)", Microsoft.Win32.RegistryValueKind.String);
            key.SetValue("URL Protocol", 0, Microsoft.Win32.RegistryValueKind.DWord);
            Microsoft.Win32.RegistryKey key2 = key.CreateSubKey("DefaultIcon");
            key2.SetValue("", "\"" + System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe") + "\"");
            key2 = key.CreateSubKey("shell");
            key2 = key2.CreateSubKey("OPEN");
            key2 = key2.CreateSubKey("command");
            key2.SetValue("", "\"" + System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe") + "\" \"%1\"");
        }

        static void SetUrlCapabilities(string urlType,string handlerName)
        {
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"Software\AMTANGEE\CRM\Capabilities\UrlAssociations");
            regKey.SetValue(urlType, handlerName);
        }

        static void SetFileCapabilities(string urlType, string handlerName)
        {
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"Software\AMTANGEE\CRM\Capabilities\FileAssociations");
            regKey.SetValue(urlType, handlerName);
        }

        static void SetCallHandler()
        {
            RegistryKey key = Registry.ClassesRoot.OpenSubKey("AMTANGEE.CallNumberUrl");
            key.SetValue("", "AMTANGEE");
            key = key.CreateSubKey("DefaultIcon");
            key.SetValue("", "\"" + System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe") + "\"");
            key = key.CreateSubKey(@"shell\open\command");
            key.SetValue("", "\"" + System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe") + "\" amtangee://call/?number=\"%1\"");
        }

        static void RegisterCall()
        {
            SetCallHandler();
            SetUrlCapabilities("tel", "AMTANGEE.CallNumberUrl");
            SetUrlCapabilities("callto", "AMTANGEE.CallNumberUrl");
            SetRegisteredApplication();
            SetDefaultCall();
            SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
        }

        static void SetDefaultCall()
        {
            RegistryKey key = Registry.ClassesRoot.OpenSubKey("tel");
            key.SetValue("AppUserModelID", "AMTANGEE.CallNumberUrl");
            key.SetValue("URL Protocol", "");
            key.SetValue("", "URL:tel Protocol");
        }

        static void SetRegisteredApplication()
        {
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(@"Software\RegisteredApplications");
            regKey.SetValue("AMTANGEE", @"SOFTWARE\AMTANGEE\CRM\Capabilities");
        }

        public static void SetAssociation(string Extension, string KeyName, string OpenWith, string OpenWithParameter, string IconIndex, string FileDescription)
        {
            try
            {
                RegistryKey key = Registry.ClassesRoot.OpenSubKey(Extension);
                string temp1 = key.GetValue(null, "").ToString();
                if (temp1.Trim().Length > 0)
                    Registry.ClassesRoot.DeleteSubKeyTree(temp1);

                Registry.ClassesRoot.DeleteSubKeyTree(Extension);
            }
            catch
            {
            }

            RegistryKey BaseKey;
            RegistryKey OpenMethod;
            RegistryKey Shell;
            RegistryKey CurrentUser;

            BaseKey = Registry.ClassesRoot.CreateSubKey(Extension);
            BaseKey.SetValue("", KeyName);

            OpenMethod = Registry.ClassesRoot.CreateSubKey(KeyName);
            OpenMethod.SetValue("", FileDescription);
            OpenMethod.CreateSubKey("DefaultIcon").SetValue("", "\"" + OpenWith + "\"," + IconIndex);
            Shell = OpenMethod.CreateSubKey("Shell");
            Shell.CreateSubKey("edit").CreateSubKey("command").SetValue("", "\"" + OpenWith + "\"" + " \"" + OpenWithParameter + "\" \"%1\"");
            Shell.CreateSubKey("open").CreateSubKey("command").SetValue("", "\"" + OpenWith + "\"" + " \"" + OpenWithParameter + "\" \"%1\"");
            BaseKey.Close();
            OpenMethod.Close();
            Shell.Close();
            try
            {
                CurrentUser = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\" + Extension, true);
                CurrentUser.DeleteSubKey("UserChoice", false);
                CurrentUser.Close();
            }
            catch
            {
            }
            // Tell explorer the file association has been changed
            SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
        }

        [System.Runtime.InteropServices.DllImport("shell32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
        
        static void RegisterFileOpen()
        {
            SetAssociation(".eml", "EML_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENEML", "4", "EML-Datei");
            SetAssociation(".msg", "MSG_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENMSG", "4", "MSG-Datei");
            SetAssociation(".vcf", "VCF_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENVCF", "5", "VCF-Datei");
            SetAssociation(".ics", "ICS_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENICS", "6", "ICS-Datei");
            SetAssociation(".vcs", "VCS_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENVCS", "6", "VCS-Datei");
        }

        static void RegisterEML()
        {
            SetAssociation(".eml", "EML_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "EML", "4", "EML-Datei");
        }

        static void RegisterSendTo()
        {
            return; // TODO: Check for compatibility for new AMTANGEE versio
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.SendTo));
            System.IO.FileInfo[] files = di.GetFiles();
            bool found = false;
            foreach (System.IO.FileInfo file in files)
            {
                if (file.Name.ToUpper().StartsWith("AMTANGEE® DOKUMENTE"))
                {
                    found = true;
                    break;
                }
            }

            if (found)
                return;
        }

        public static bool IsUserAdministrator()
        {
            bool isAdmin;
            WindowsIdentity user = null;
            try
            {
                user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException ex)
            {
                isAdmin = false;
            }
            catch (Exception ex)
            {
                isAdmin = false;
            }
            finally
            {
                if (user != null)
                    user.Dispose();
            }
            return isAdmin;
        }

        [STAThread]
        static void Main(string[] param)
        {
            _param2 = param;
            if (param != null && param.Length > 0)
            {
                _param = param[0];
                if (param.Length > 1)
                    _param_2 = param[1];
                else
                    _param_2 = "";
                if (_param.ToUpper() == "CLOSE")
                {
                    System.Diagnostics.Process.GetCurrentProcess().Kill();
                }

                if (_param.ToUpper() == "SETDB")
                {
                    if (_param2.Length > 1 && _param2[1].Trim().Length > 0)
                    {
                        Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\AMTANGEE\RemoteUrl");
                        if (regKey != null)
                        {
                            if (_param2[1].Trim() == "0")
                            {
                                try
                                {
                                    regKey.SetValue("DefaultDB", "0");
                                    regKey.DeleteValue("DefaultDB");
                                }
                                catch { }
                            }
                            else
                                regKey.SetValue("DefaultDB", _param2[1].Trim(), RegistryValueKind.String);
                        }

                        if (_param2[1].Trim() == "0")
                            MessageBox.Show("Standard-Datenbank für FileViewer wurde auf Standard-Einstellungen zurückgesetzt!");
                        else
                            MessageBox.Show("Standard-Datenbank für FileViewer wurde auf '" + _param2[1].Trim() + "' gesetzt!");
                    }
                    return;
                }

                if (_param == "REG_ALL" || _param == "REG_ALL_ADMIN")
                {
                    try
                    {
                        if (!IsUserAdministrator() && _param != "REG_ALL_ADMIN")
                        {
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            process.StartInfo.FileName = System.Reflection.Assembly.GetEntryAssembly().Location;
                            process.StartInfo.Arguments = "REG_ALL_ADMIN";
                            process.StartInfo.Verb = "runas";
                            process.Start();
                            return;
                        }
                        RegisterUrl();
                        RegisterSendTo();
                        RegisterEML();
                        RegisterFileOpen();
                        RegisterCall();
                    }
                    catch
                    {
                    }
                    return;
                }

                if (_param == "REG_URL" || _param == "REG_URL_ADMIN")
                {
                    try
                    {

                        if (!IsUserAdministrator() && _param != "REG_URL_ADMIN")
                        {
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            process.StartInfo.FileName = System.Reflection.Assembly.GetEntryAssembly().Location;
                            process.StartInfo.Arguments = "REG_URL_ADMIN";
                            process.StartInfo.Verb = "runas";
                            process.Start();
                            return;
                        }
                        RegisterUrl();
                    }
                    catch
                    {
                    }

                    MessageBox.Show("Registrierung erfolgreich!");
                    return;

                }
                if (_param == "REG_CALL" || _param == "REG_CALL_ADMIN")
                {
                    try
                    {

                        if (!IsUserAdministrator() && _param != "REG_CALL_ADMIN")
                        {
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            process.StartInfo.FileName = System.Reflection.Assembly.GetEntryAssembly().Location;
                            process.StartInfo.Arguments = "REG_CALL_ADMIN";
                            process.StartInfo.Verb = "runas";
                            process.Start();
                            return;
                        }
                        RegisterCall();
                    }
                    catch
                    {
                    }

                    MessageBox.Show("Registrierung erfolgreich!");
                    return;

                }
                if (_param == "REG_WITHOUT_PARAM_ADMIN")
                {
                    try
                    {
                        RegisterFileOpen();
                    }
                    catch
                    {
                    }

                    return;
                }


                if (_param == "REG_EML" || _param == "REG_EML_ADMIN")
                {
                    try
                    {
                        if (!IsUserAdministrator() && _param != "REG_EML_ADMIN")
                        {
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            process.StartInfo.FileName = System.Reflection.Assembly.GetEntryAssembly().Location;
                            process.StartInfo.Arguments = "REG_EML_ADMIN";
                            process.StartInfo.Verb = "runas";
                            process.Start();
                            return;
                        }
                        RegisterEML();
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show("Fehler beim Registrieren! Fehler: \r\n" + exc.Message);
                    }

                    MessageBox.Show("Registrierung erfolgreich!");
                    return;

                }

                if (_param == "REG_SENDTO")
                {
                    try
                    {
                        RegisterSendTo();
                    }
                    catch
                    {
                    }

                    MessageBox.Show("Registrierung erfolgreich!");
                    return;
                }


                try
                {
                    if (System.IO.File.Exists(_param))
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(_param);
                        bool found = false;
                        switch (fi.Extension.Replace(".", "").ToUpper())
                        {
                            case "EML":
                                SetAssociation(".eml", "EML_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENEML", "4", "EML-Datei");
                                found = true;
                                break;

                            case "MSG":
                                SetAssociation(".msg", "MSG_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENMSG", "4", "MSG-Datei");
                                found = true;
                                break;

                            case "VCF":
                                SetAssociation(".vcf", "VCF_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENVCF", "5", "VCF-Datei");
                                found = true;
                                break;

                            case "ICS":
                                SetAssociation(".ics", "ICS_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENICS", "6", "ICS-Datei");
                                found = true;
                                break;

                            case "VCS":
                                SetAssociation(".vcs", "VCS_File", System.IO.Path.Combine(AMTANGEE.SDK.Global.AMTANGEEDirectory, "FileViewer.exe"), "OPENVCS", "6", "VCS-Datei");
                                found = true;
                                break;
                        }

                        if (found)
                        {
                            System.Diagnostics.Process.Start(_param);
                            return;
                        }

                    }
                }
                catch
                {
                }
            }
            else
            {
                if (!IsUserAdministrator())
                {
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = System.Reflection.Assembly.GetEntryAssembly().Location;
                    process.StartInfo.Arguments = "REG_WITHOUT_PARAM_ADMIN";
                    process.StartInfo.Verb = "runas";
                    process.Start();
                    return;
                }
                RegisterFileOpen();
                return;
            }
            SingleInstanceController controller = new SingleInstanceController();
            controller.Run(param);
        }

        


        static string GetOption(string parameter,string option)
        {
            if (parameter.ToLower().StartsWith("amtangee://"))
                parameter = parameter.Substring(11);
            string returnString = "";
            try
            {
                if (parameter.ToLower().IndexOf(option.ToLower()) >= 0)
                {
                    int pos = parameter.ToLower().IndexOf(option.ToLower());
                    if (pos >= 0)
                    {
                        pos = pos + option.Length + 1;
                        int pos2 = parameter.ToLower().IndexOf("?",pos);
                        if (parameter[parameter.Length - 1] == '/')
                            parameter = parameter.Remove(parameter.Length - 1);

                        if (pos2 >= 0)
                        {
                            returnString = parameter.Substring(pos, pos2);
                        }
                        else
                        {
                            pos2 = parameter.IndexOf("&", pos);
                            if (pos2 >= 0)
                            {
                                returnString = parameter.Substring(pos, pos2 - pos);
                            }
                            else
                            {
                                returnString = parameter.Substring(pos);
                            }
                        }
                    }
                }
            }
            catch
            {
            }
            returnString = System.Web.HttpUtility.UrlDecode(returnString);
            return returnString;
        }
    }
}