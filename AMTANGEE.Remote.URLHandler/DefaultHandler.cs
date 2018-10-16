using System;
using System.Collections.Generic;
using System.Text;

namespace AMTANGEE
{
    public partial class Remote
    {
        public class DefaultHandler
        {
            class DummyHost : AMTANGEE.SDK.IHost
            {
                public void AddWindow(string id)
                {

                }

                public bool CompletelyLoaded()
                {
                    return true;
                }

                public void DotNetMenuClosed(object menu)
                {

                }

                public bool HasParameter(string parameter)
                {
                    return false;
                }

                public object LoadedModules()
                {
                    return null;
                }

                public bool ModuleLoaded(string module)
                {
                    return true;
                }

                public object NewContact(string category, string name, string name2, string street, string country, string zip, string city, string areaCode, string phoneNumber, string phoneNumberType)
                {
                    return null;
                }

                public void ProcessEvent(string module, string theEvent)
                {

                }

                public void RemoveWindow(string id)
                {

                }

                public object SearchContacts(string caption, string buttonCaption, bool allowMultiSelect)
                {
                    return null;
                }

                public object SearchObject(string caption, string buttonCaption, bool searchForContacts, bool searchForContactPersons, bool searchForDocuments, bool searchForTasks, bool searchForAppointments, bool searchForOpportunities, bool searchForOffers, bool searchForOrders, bool searchForInvoices, bool searchForProjects, bool searchForEmails, bool searchForFaxes, bool searchForCalls, bool searchForShortMessages, bool searchForNotifications, bool searchForSmses)
                {
                    return null;
                }

                public object SelectProject(string contact)
                {
                    return null;
                }


                public void SearchContactAndAssign(string caption, string buttonCaption, string guid, bool modal)
                {
                   
                }

                public object SearchObjects(string caption, string buttonCaption, bool searchForContacts, bool searchForContactPersons, bool searchForDocuments, bool searchForTasks, bool searchForAppointments, bool searchForOpportunities, bool searchForOffers, bool searchForOrders, bool searchForInvoices, bool searchForProjects, bool searchForEmails, bool searchForFaxes, bool searchForCalls, bool searchForShortMessages, bool searchForNotifications, bool searchForSmses)
                {
                    return null;
                }

                public void AddForm(string id)
                {
                  
                }

                public void RemoveForm(string id)
                {
                   
                }
            }

            #region SMS
            public class SMS : IRemoteHandler
            {
                public void Exec(Paramters parameter)
                {
                    if (parameter["number"].Length > 0 && parameter["text"].Length > 0)
                    {
                        AMTANGEE.SDK.Messages.Sms.Send(AMTANGEE.SDK.Global.CurrentUser, parameter["number"], parameter["text"]);
                    }
                }

                public string ModuleName
                {
                    get { return "SMS"; }
                }
            }
            #endregion

            #region Call

            public class Call : IRemoteHandler
            {
                public string ModuleName
                {
                    get { return "CALL"; }
                }

                public void Exec(Paramters parameter)
                {
                    if (!Tools.CheckForAMTANGEE())
                        return;
                    if (parameter.GetValue("number").Length > 0)
                    { 
                        AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "INTEGRA:NEWCALL:" + AMTANGEE.SDK.Contacts.PhoneNumbers.Clean(parameter["number"].Replace("tel:","").Replace("callto:","")) + ",{00000000-0000-0000-0000-000000000000}", true);
                        return;
                    }

                    if (parameter["contact"].Length > 0)
                    {
                        if (parameter["contact"].Length == 36)
                        {
                            AMTANGEE.SDK.Contacts.Contact c = new AMTANGEE.SDK.Contacts.Contact(new Guid(parameter["contact"]));
                            c.Call();
                        }
                        return;
                    }
                }
            }

            #endregion

            #region Projects
            public class Project : IRemoteHandler
            {
                AMTANGEE.SDK.Projects.Project GetProject(Paramters parameter)
                {
                    AMTANGEE.SDK.Projects.Project project = null;
                    if (parameter["guid"].Length > 0)
                    {
                        try
                        {
                            project = new SDK.Projects.Project(new Guid(parameter["guid"]));

                            if (!project.ExistsAndLoadedAndRights)
                                project = null;
                        }
                        catch
                        {
                        }
                    }

                    if (project == null)
                    {
                        if (parameter["projectno"].Trim().Length > 0)
                        {

                            AMTANGEE.SDK.Projects.Projects projects = new AMTANGEE.SDK.Projects.Projects(AMTANGEE.SDK.Users.Users.Public, SDK.Projects.Projects.Kinds.All);
                            string projectNO = parameter["projectno"].Trim();
                            foreach (AMTANGEE.SDK.Projects.Project pro in projects)
                            {
                                if (pro.ProjectNo.Trim().ToUpper() == projectNO.Trim().ToUpper())
                                {
                                    project = pro;
                                    break;
                                }
    
                            }
                          

                        }
                    }
                    


                    return project;
                }

                public void Exec(Paramters parameter)
                {

                    if (parameter["action"].ToLower() == "show")
                    {
                        if (!Tools.CheckForAMTANGEE(false))
                            return;

                        AMTANGEE.SDK.Projects.Project project = GetProject(parameter);

                        if (project != null)
                            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "*AMTANGEE.Modules.Projects.dll:EDIT:{" + project.Guid.ToString() + "}");
                    }


                    if (parameter["action"].ToLower() == "create")
                    {
                        AMTANGEE.SDK.Projects.Project project = new AMTANGEE.SDK.Projects.Project(AMTANGEE.SDK.Global.CurrentUser);
                        project.Description = parameter["description"].Trim();
                        project.ProjectNo = parameter["projectnumber"].Trim();
                        project.Start = DateTime.Now;
                        AMTANGEE.SDK.Contacts.ContactBase cb = null;
                        if (parameter["guid"].Length > 0)
                        {
                            try
                            {
                                cb = AMTANGEE.SDK.Contacts.ContactBase.Create(new Guid(parameter["guid"]));
                                if (!cb.ExistsAndLoadedAndRights)
                                    cb = null;
                            }
                            catch
                            {
                            }
                        }

                        if (cb == null)
                        {
                            if (parameter["vendor"].Length > 0 || parameter["customer"].Length > 0)
                            {
                                string custorvendid = "";
                                if (parameter["customer"].Length > 0)
                                {
                                    custorvendid = parameter["customer"];
                                }
                                else
                                {
                                    if (parameter["vendor"].Length > 0)
                                        custorvendid = parameter["vendor"];
                                }
                                cb = Tools.GetContactByCustomerOrVendorNo(custorvendid, parameter["type"]);
                            }
                            
                        }
                        if (cb == null)
                        {
                            if (parameter["source"].Length > 0)
                            {
                                cb = Tools.GetContactBySource(parameter["source"], parameter["type"]);
                            }

                        }

                        if (cb == null)
                        {
                            if (parameter["phonenumber"].Length > 0)
                            {
                                if (parameter["type"].ToLower() == "person")
                                {
                                    cb = AMTANGEE.SDK.Contacts.Contacts.SearchByPhoneNumber(parameter["phonenumber"]);
                                    if (cb != null)
                                        if (!(cb is AMTANGEE.SDK.Contacts.ContactPerson))
                                            cb = null;
                                }
                                else
                                {
                                    cb = AMTANGEE.SDK.Contacts.Contacts.SearchByPhoneNumber(parameter["phonenumber"]);
                                }
                            }
                        }

                        if (cb == null)
                        {
                            if (parameter["emailadress"].Length > 0)
                            {
                                if (parameter["type"].ToLower() == "person")
                                {
                                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchByEmailAddress(AMTANGEE.SDK.Contacts.SearchKinds.Contains, parameter["emailadress"]);
                                    if (contacts.Count > 0)
                                    {
                                        foreach (AMTANGEE.SDK.Contacts.ContactBase contactBase in contacts)
                                        {
                                            if (contactBase is AMTANGEE.SDK.Contacts.ContactPerson)
                                            {
                                                cb = contactBase;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchByEmailAddress(AMTANGEE.SDK.Contacts.SearchKinds.Contains, parameter["emailadress"]);
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
                            }
                        }

                        if (cb != null)
                        {
                            project.Contact = cb;
                            project.Contacts.Add(cb);
                        }

                        project.Save();
                    }


                }

                public string ModuleName
                {
                    get { return "PROJECT"; }
                }
            }
            #endregion

            #region General
            public class General : IRemoteHandler
            {
                public void Exec(Paramters parameter)
                {
                    if (parameter["action"].ToLower() == "deletereference")
                    {
                        Tools.Settings[parameter["[REF]"].ToUpper().Trim()] = "";
                        return;
                    }
                }

                public string ModuleName
                {
                    get { return "GENERAL"; }
                }
            }

            #endregion

            #region Document
            public class Document : IRemoteHandler
            {
                public void Exec(Paramters parameter)
                {
                    if (parameter["action"].ToLower() == "delete")
                    {
                        try
                        {
                            AMTANGEE.SDK.Documents.Document doc = new AMTANGEE.SDK.Documents.Document(new Guid(parameter["[REFGUID]"]));
                            doc.Delete();
                            doc.Save();
                        }
                        catch { }
                    }
                    if (parameter["action"].ToLower() == "new")
                    {
                        if (parameter["file"].Trim().Length > 0)
                        {
                            string filename = parameter["file"].Trim();
                            try
                            {
                                if (System.IO.File.Exists(filename))
                                {
                                    AMTANGEE.SDK.Documents.Category catForDocument = null;
                                    if (parameter["category"].Trim().Length > 0)
                                    {
                                        string catName = parameter["category"];
                                        try
                                        {
                                            catForDocument = new AMTANGEE.SDK.Documents.Category(new Guid(catName));
                                        }
                                        catch
                                        {
                                            catForDocument = null;
                                        }

                                        if (catForDocument == null)
                                        {
                                            AMTANGEE.SDK.Documents.Categories categories = new AMTANGEE.SDK.Documents.Categories(AMTANGEE.SDK.Global.CurrentUser, false);
                                            foreach (AMTANGEE.SDK.Documents.Category cat in categories)
                                            {
                                                catForDocument = (AMTANGEE.SDK.Documents.Category)Tools.GetCategoryByFullName(cat, catName);
                                                if (catForDocument != null)
                                                    break;
                                            }
                                        }
                                    }

                                    if (catForDocument == null)
                                    {
                                        AMTANGEE.Forms.SelectCategoryForm scf = new AMTANGEE.Forms.SelectCategoryForm(AMTANGEE.Controls.CategoriesTree.Kinds.Documents, AMTANGEE.SDK.CategoryRights.Add, "Zielkategorie wählen", "Importieren");
                                        if (scf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                        {
                                            catForDocument = (AMTANGEE.SDK.Documents.Category)scf.SelectedCategory;
                                        }
                                    }
                                    if (catForDocument != null)
                                    {
                                        AMTANGEE.SDK.Documents.Document doc = new AMTANGEE.SDK.Documents.Document(catForDocument);
                                        doc.LoadFromFile(filename);
                                        doc.Description = System.IO.Path.GetFileNameWithoutExtension(filename);
                                        doc.DateTime = DateTime.Now;
                                        doc.FileType = System.IO.Path.GetExtension(filename);
                                        doc.Save();
                                        if (parameter["[REF]"].Length > 0)
                                            Tools.Settings[parameter["[REF]"]] = doc.Guid.ToString();
                                        return;
                                    }
                                }
                            }
                            catch
                            {
                            }
                        }
                    }
                }

                public string ModuleName
                {
                    get { return "DOCUMENT"; }
                }
            }

            #endregion

            #region Appointment
            public class Appointment : IRemoteHandler
            {
                public void Exec(Paramters parameter)
                {

                    AMTANGEE.SDK.Appointments.Appointment appointment = null;
                    if (parameter["action"].ToLower() == "addlink")
                    {
                        try
                        {
                            if (parameter["[REFGUID]"].Length > 0)
                            {

                                appointment = new AMTANGEE.SDK.Appointments.Appointment(new Guid(parameter["[REFGUID]"]));
                                if (parameter["type"].ToLower() == "document")
                                {
                                    AMTANGEE.SDK.Documents.Document doc = new AMTANGEE.SDK.Documents.Document(new Guid(Tools.Settings[parameter["documentreference"].ToUpper().Trim()]));
                                    appointment.Links.Add(doc);
                                    appointment.Save();
                                }

                            }
                        }
                        catch { }
                    }

                    if (parameter["action"].ToLower() == "addparticipant")
                    {
                        try
                        {
                            if (parameter["[REFGUID]"].Length > 0)
                            {

                                appointment = new AMTANGEE.SDK.Appointments.Appointment(new Guid(parameter["[REFGUID]"]));

                                if (parameter["type"].ToLower() == "resource")
                                {
                                    AMTANGEE.SDK.Appointments.Resources resources = new AMTANGEE.SDK.Appointments.Resources(AMTANGEE.SDK.Global.CurrentUser.Location);
                                    string resourceName = parameter["name"].ToLower();
                                    AMTANGEE.SDK.Appointments.Resource resource = null;
                                    foreach (AMTANGEE.SDK.Appointments.Resource res in resources)
                                    {
                                        if (res.Name.ToUpper() == resourceName.ToUpper())
                                        {
                                            resource = res;
                                            break;
                                        }
                                    }

                                    if (resource != null)
                                    {
                                        appointment.Participants.Add(resource);
                                        appointment.Save();
                                    }


                                }
                                else
                                {
                                    if (parameter["type"].ToLower() == "user")
                                    {
                                        AMTANGEE.SDK.Users.UserBase ub2 = AMTANGEE.SDK.Global.CurrentUser;
                                        if (parameter["name"].Trim().Length > 0)
                                        {
                                            bool foundUser = false;
                                            try
                                            {
                                                AMTANGEE.SDK.Users.User user = new AMTANGEE.SDK.Users.User(new Guid(parameter["name"].Trim()));
                                                if (user.ExistsAndLoadedAndRights)
                                                {
                                                    foundUser = true;
                                                    ub2 = user;
                                                }
                                            }
                                            catch
                                            {
                                            }


                                            if (!foundUser)
                                            {
                                                string userName = parameter["name"].Trim().ToUpper();
                                                AMTANGEE.SDK.Users.Users users = new AMTANGEE.SDK.Users.Users(AMTANGEE.SDK.Global.CurrentUser.Location);


                                                foreach (AMTANGEE.SDK.Users.User user in users)
                                                {
                                                    if (user.Name.ToUpper() == userName)
                                                    {
                                                        foundUser = true;
                                                        ub2 = user;
                                                        break;
                                                    }
                                                }

                                                if (!foundUser)
                                                {
                                                    foreach (AMTANGEE.SDK.Users.User user in users)
                                                    {
                                                        if (user.DisplayName.ToUpper() == userName)
                                                        {
                                                            foundUser = true;
                                                            ub2 = user;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (ub2 != null)
                                        {
                                            appointment.Participants.Add(ub2);
                                            appointment.Save();
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        AMTANGEE.SDK.Contacts.ContactBase cb = null;
                                        if (parameter["guid"].Length > 0)
                                        {
                                            try
                                            {
                                                cb = AMTANGEE.SDK.Contacts.ContactBase.Create(new Guid(parameter["guid"]));
                                                if (!cb.ExistsAndLoadedAndRights)
                                                    cb = null;
                                            }
                                            catch
                                            {
                                            }
                                        }

                                        if (cb == null)
                                        {
                                            if (parameter["vendor"].Length > 0 || parameter["customer"].Length > 0)
                                            {
                                                string custorvendid = "";
                                                if (parameter["customer"].Length > 0)
                                                {
                                                    custorvendid = parameter["customer"];
                                                }
                                                else
                                                {
                                                    if (parameter["vendor"].Length > 0)
                                                        custorvendid = parameter["vendor"];
                                                }
                                               
                                                    cb = Tools.GetContactByCustomerOrVendorNo(custorvendid, parameter["type"]);
                                                
                                            }
                                        }
                                        if (cb == null)
                                        {
                                            if (parameter["source"].Length > 0)
                                            {
                                                cb = Tools.GetContactBySource(parameter["source"], parameter["type"]);
                                            }

                                        }
                                        if (cb == null)
                                        {
                                            if (parameter["phonenumber"].Length > 0)
                                            {
                                                if (parameter["type"].ToLower() == "person")
                                                {
                                                    cb = AMTANGEE.SDK.Contacts.Contacts.SearchByPhoneNumber(parameter["phonenumber"]);
                                                    if (cb != null)
                                                        if (!(cb is AMTANGEE.SDK.Contacts.ContactPerson))
                                                            cb = null;
                                                }
                                                else
                                                {
                                                    cb = AMTANGEE.SDK.Contacts.Contacts.SearchByPhoneNumber(parameter["phonenumber"]);
                                                }
                                            }
                                        }

                                        if (cb == null)
                                        {
                                            if (parameter["emailadress"].Length > 0)
                                            {
                                                if (parameter["type"].ToLower() == "person")
                                                {
                                                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchByEmailAddress(AMTANGEE.SDK.Contacts.SearchKinds.Contains, parameter["emailadress"]);
                                                    if (contacts.Count > 0)
                                                    {
                                                        foreach (AMTANGEE.SDK.Contacts.ContactBase contactBase in contacts)
                                                        {
                                                            if (contactBase is AMTANGEE.SDK.Contacts.ContactPerson)
                                                            {
                                                                cb = contactBase;
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchByEmailAddress(AMTANGEE.SDK.Contacts.SearchKinds.Contains, parameter["emailadress"]);
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
                                            }
                                        }

                                        if (cb != null)
                                        {
                                            appointment.Participants.Add(cb);
                                            appointment.Save();
                                            return;
                                        }
                                    }
                                }

                                return;
                            }
                        }
                        catch
                        {
                        }
                    }
                    if (parameter["action"].ToLower() == "delete")
                    {
                        try
                        {
                            if (parameter["[REFGUID]"].Length > 0)
                            {
                                appointment = new AMTANGEE.SDK.Appointments.Appointment(new Guid(parameter["[REFGUID]"]));
                                appointment.Delete();
                                appointment.Save();
                                return;
                            }
                        }
                        catch
                        {
                        }
                    }

                    AMTANGEE.SDK.Users.UserBase ub = AMTANGEE.SDK.Global.CurrentUser;
                    if (parameter["user"].Trim().Length > 0)
                    {
                        bool foundUser = false;
                        try
                        {
                            AMTANGEE.SDK.Users.User user = new AMTANGEE.SDK.Users.User(new Guid(parameter["user"].Trim()));
                            if (user.ExistsAndLoadedAndRights)
                            {
                                foundUser = true;
                                ub = user;
                            }
                        }
                        catch
                        {
                        }


                        if (!foundUser)
                        {
                            string userName = parameter["user"].Trim().ToUpper();
                            AMTANGEE.SDK.Users.Users users = new AMTANGEE.SDK.Users.Users(AMTANGEE.SDK.Global.CurrentUser.Location);


                            foreach (AMTANGEE.SDK.Users.User user in users)
                            {
                                if (user.Name.ToUpper() == userName)
                                {
                                    foundUser = true;
                                    ub = user;
                                    break;
                                }
                            }

                            if (!foundUser)
                            {
                                foreach (AMTANGEE.SDK.Users.User user in users)
                                {
                                    if (user.DisplayName.ToUpper() == userName)
                                    {
                                        foundUser = true;
                                        ub = user;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    if (parameter["action"].ToLower() == "modify")
                    {
                        try
                        {
                            if (parameter["[REFGUID]"].Length > 0)
                                appointment = new AMTANGEE.SDK.Appointments.Appointment(new Guid(parameter["[REFGUID]"]));

                            if (parameter["subject"].Trim().Length > 0)
                                appointment.Description = parameter["subject"].Trim();

                            if (parameter["start"].Trim().Length > 0)
                            {
                                try
                                {
                                    appointment.Start = DateTime.Parse(parameter["start"].Trim());
                                }
                                catch
                                {
                                }
                            }

                            if (parameter["end"].Trim().Length > 0)
                            {
                                try
                                {
                                    appointment.End = DateTime.Parse(parameter["end"].Trim());
                                }
                                catch
                                {
                                }
                            }

                            if (parameter["user"].Trim().Length > 0)
                                appointment.AssignedTo = ub;

                            if (parameter["notes"].Trim().Length > 0)
                                appointment.Notes = parameter["notes"].Trim();

                            if (parameter["category"].Trim().Length > 0)
                            {
                                string catName = parameter["category"].Trim().ToUpper();
                                AMTANGEE.SDK.Appointments.Categories categories = new AMTANGEE.SDK.Appointments.Categories(ub);
                                foreach (AMTANGEE.SDK.Appointments.Category category in categories)
                                {
                                    if (category.Name.ToUpper() == catName)
                                    {
                                        appointment.Category = category;
                                        break;
                                    }
                                }
                            }

                            appointment.Save();

                            return;
                        }
                        catch
                        {
                        }
                    }

                    if (parameter["action"].ToLower() == "new")
                    {
                        try
                        {
                            if (parameter["subject"].Trim().Length > 0)
                            {
                                try
                                {
                                    appointment = new AMTANGEE.SDK.Appointments.Appointment(ub);
                                    appointment.Description = parameter["subject"].Trim();
                                    if (parameter["start"].Trim().Length > 0)
                                    {
                                        try
                                        {
                                            appointment.Start = DateTime.Parse(parameter["start"].Trim());
                                        }
                                        catch
                                        {
                                        }
                                    }

                                    if (parameter["end"].Trim().Length > 0)
                                    {
                                        try
                                        {
                                            appointment.End = DateTime.Parse(parameter["end"].Trim());
                                        }
                                        catch
                                        {
                                        }
                                    }

                                    appointment.Notes = parameter["notes"].Trim();

                                    if (parameter["category"].Trim().Length > 0)
                                    {
                                        string catName = parameter["category"].Trim().ToUpper();
                                        AMTANGEE.SDK.Appointments.Categories categories = new AMTANGEE.SDK.Appointments.Categories(ub);
                                        foreach (AMTANGEE.SDK.Appointments.Category category in categories)
                                        {
                                            if (category.Name.ToUpper() == catName)
                                            {
                                                appointment.Category = category;
                                                break;
                                            }
                                        }
                                    }

                                    appointment.Save();
                                    if (parameter["[REF]"].Length > 0)
                                        Tools.Settings[parameter["[REF]"]] = appointment.Guid.ToString();

                                    return;
                                }
                                catch
                                {
                                }
                            }

                        }
                        catch
                        {
                        }
                    }
                }

                public string ModuleName
                {
                    get { return "APPOINTMENT"; }
                }
            }

            #endregion

            #region Contact
            public class Contact : IRemoteHandler
            {

                AMTANGEE.SDK.Contacts.ContactBase GetContact(Paramters parameter)
                {
                    AMTANGEE.SDK.Contacts.ContactBase cb = null;
                    if (parameter["guid"].Length > 0)
                    {
                        try
                        {
                            cb = AMTANGEE.SDK.Contacts.ContactBase.Create(new Guid(parameter["guid"]));
                            if (!cb.ExistsAndLoadedAndRights)
                                cb = null;
                        }
                        catch
                        {
                        }
                    }

                    if (cb == null)
                    {
                        if (parameter["vendor"].Length > 0 || parameter["customer"].Length > 0)
                        {
                            string custorvendid = "";
                            if (parameter["customer"].Length > 0)
                            {
                                custorvendid = parameter["customer"];
                            }
                            else
                            {
                                if (parameter["vendor"].Length > 0)
                                    custorvendid = parameter["vendor"];
                            }
                            cb = Tools.GetContactByCustomerOrVendorNo(custorvendid, parameter["type"]);

                        }
                    }
                    if (cb == null)
                    {
                        if (parameter["source"].Length > 0)
                        {
                            cb = Tools.GetContactBySource(parameter["source"], parameter["type"]);
                        }

                    }
                    if (cb == null)
                    {
                        if (parameter["phonenumber"].Length > 0)
                        {
                            if (parameter["type"].ToLower() == "person")
                            {
                                cb = AMTANGEE.SDK.Contacts.Contacts.SearchByPhoneNumber(parameter["phonenumber"]);
                                if (cb != null)
                                    if (!(cb is AMTANGEE.SDK.Contacts.ContactPerson))
                                        cb = null;
                            }
                            else
                            {
                                cb = AMTANGEE.SDK.Contacts.Contacts.SearchByPhoneNumber(parameter["phonenumber"]);
                            }
                        }
                    }

                    if (cb == null)
                    {
                        if (parameter["emailadress"].Length > 0)
                        {
                            if (parameter["type"].ToLower() == "person")
                            {
                                AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchByEmailAddress(AMTANGEE.SDK.Contacts.SearchKinds.Contains, parameter["emailadress"]);
                                if (contacts.Count > 0)
                                {
                                    foreach (AMTANGEE.SDK.Contacts.ContactBase contactBase in contacts)
                                    {
                                        if (contactBase is AMTANGEE.SDK.Contacts.ContactPerson)
                                        {
                                            cb = contactBase;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                AMTANGEE.SDK.Contacts.Contacts contacts = AMTANGEE.SDK.Contacts.Contacts.SearchByEmailAddress(AMTANGEE.SDK.Contacts.SearchKinds.Contains, parameter["emailadress"]);
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
                        }
                    }




                    return cb;
                }

                public void Exec(Paramters parameter)
                {
                    bool force = false;
                    if (parameter["donotcheck"].ToUpper() == "TRUE" || parameter["donotcheck"].ToUpper() == "1")
                        force = true;

                    AMTANGEE.SDK.Contacts.Contact c = null;
                    if (parameter["action"].ToLower() == "new")
                    {

                        AMTANGEE.SDK.Contacts.Category catForContact = null;
                        if (parameter["category"].Trim().Length > 0)
                        {
                            string catName = parameter["category"];
                            try
                            {
                                catForContact = new AMTANGEE.SDK.Contacts.Category(new Guid(catName));
                            }
                            catch
                            {
                                catForContact = null;
                            }

                            if (catForContact == null)
                            {
                                AMTANGEE.SDK.Contacts.Categories categories = new AMTANGEE.SDK.Contacts.Categories(AMTANGEE.SDK.Global.CurrentUser, false);
                                foreach (AMTANGEE.SDK.Contacts.Category cat in categories)
                                {
                                    catForContact = (AMTANGEE.SDK.Contacts.Category)Tools.GetCategoryByFullName(cat, catName);
                                    if (catForContact != null)
                                        break;
                                }
                            }
                        }

                        if (catForContact == null)
                        {
                            AMTANGEE.Forms.SelectCategoryForm scf = new AMTANGEE.Forms.SelectCategoryForm(AMTANGEE.Controls.CategoriesTree.Kinds.Contacts, AMTANGEE.SDK.CategoryRights.Add, "Zielkategorie wählen", "Importieren");
                            if (scf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                catForContact = (AMTANGEE.SDK.Contacts.Category)scf.SelectedCategory;
                            }
                        }
                        if (catForContact != null)
                        {
                            //Kontakt anlegen...

                        }

                    }

                    if (parameter["action"].ToLower() == "simulate_call_inbound")
                    {
                        if (!Tools.CheckForAMTANGEE(force))
                            return;

                        AMTANGEE.SDK.Contacts.ContactBase cb = GetContact(parameter);

                        if (cb != null)
                            cb.SimulateCall(true);
                    }

                    if (parameter["action"].ToLower() == "simulate_call_outbound")
                    {
                        if (!Tools.CheckForAMTANGEE(force))
                            return;

                        AMTANGEE.SDK.Contacts.ContactBase cb = GetContact(parameter);

                        if (cb != null)
                            cb.SimulateCall(false);
                    }

                    if (parameter["action"].ToLower() == "call")
                    {
                        if (!Tools.CheckForAMTANGEE(force))
                            return;

                        AMTANGEE.SDK.Contacts.ContactBase cb = GetContact(parameter);

                        if (cb != null)
                            cb.Call();
                    }

                    if (parameter["action"].ToLower() == "show")
                    {
                        if (!Tools.CheckForAMTANGEE(force))
                            return;

                        AMTANGEE.SDK.Contacts.ContactBase cb = GetContact(parameter);
                        

                        if (cb != null)
                            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "*AMTANGEE.Modules.Contacts.dll:EDIT:{" + cb.Guid.ToString() + "}");
                    }

                    if (parameter["action"].ToLower() == "deletehistoryentry")
                    {
                        try
                        {
                            if (parameter["[REFGUID]"].Length > 0)
                            {
                                AMTANGEE.SDK.Contacts.HistoryEntry he = new AMTANGEE.SDK.Contacts.HistoryEntry(new Guid(parameter["[REFGUID]"]));
                                he.Delete();
                                he.Save();
                                return;
                            }
                        }
                        catch
                        {
                        }
                    }
                    if (parameter["action"].ToLower() == "modifyhistoryentry")
                    {
                        if (parameter["[REFGUID]"].Length > 0)
                        {
                            AMTANGEE.SDK.Contacts.HistoryEntry he = new AMTANGEE.SDK.Contacts.HistoryEntry(new Guid(parameter["[REFGUID]"]));

                            try
                            {
                                if (parameter["subject"].Trim().Length > 0)
                                    he.Description = parameter["subject"].Trim();

                                if (parameter["datetime"].Trim().Length > 0)
                                    he.DateTime = DateTime.Parse(parameter["datetime"].Trim());

                                if (parameter["notes"].Trim().Length > 0)
                                    he.Notes = parameter["notes"].Trim();

                                if (parameter["historytype"].Length > 0)
                                {
                                    switch (parameter["historytype"].ToUpper().Trim())
                                    {
                                        case "LETTER":
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Letter;
                                            break;
                                        case "CALL":
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Call;
                                            break;
                                        case "EMAIL":
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Email;
                                            break;
                                        case "FAX":
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Fax;
                                            break;
                                        case "NOTE":
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Note;
                                            break;
                                        case "SMS":
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Sms;
                                            break;
                                        default:
                                            he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.General;
                                            break;

                                    }
                                }

                                he.Save();

                                return;
                            }
                            catch
                            {

                            }
                        }
                    }
                    if (parameter["action"].ToLower() == "addhistoryentry")
                    {
                        AMTANGEE.SDK.Contacts.ContactBase cb = GetContact(parameter);
                        

                        if (cb != null)
                        {
                            if (parameter["subject"].Length > 0 && parameter["datetime"].Length > 0)
                            {
                                AMTANGEE.SDK.Contacts.HistoryEntry he = new AMTANGEE.SDK.Contacts.HistoryEntry(cb);

                                try
                                {
                                    he.Description = parameter["subject"].Trim();
                                    he.DateTime = DateTime.Parse(parameter["datetime"].Trim());
                                    he.Notes = parameter["notes"].Trim();
                                    if (parameter["historytype"].Length > 0)
                                    {
                                        switch (parameter["historytype"].ToUpper().Trim())
                                        {
                                            case "LETTER":
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Letter;
                                                break;
                                            case "CALL":
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Call;
                                                break;
                                            case "EMAIL":
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Email;
                                                break;
                                            case "FAX":
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Fax;
                                                break;
                                            case "NOTE":
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Note;
                                                break;
                                            case "SMS":
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.Sms;
                                                break;
                                            default:
                                                he.Kind = AMTANGEE.SDK.Contacts.HistoryEntry.Kinds.General;
                                                break;

                                        }
                                    }

                                    he.Save();
                                    
                                    if (parameter["[REFID]"].Length > 0)
                                        Tools.Settings[parameter["[REFID]"]] = he.Guid.ToString();

                                    return;
                                }
                                catch
                                {

                                }
                            }
                        }
                    }
                }

                public string ModuleName
                {
                    get { return "CONTACT"; }
                }
            }

            #endregion

            #region Email
            public class Email : AMTANGEE.Remote.IRemoteHandler
            {
                public void Exec(Paramters parameter)
                {
                    //if (parameter["action"] == "new")
                    //{
                    //    if (parameter["subject"].Length > 0)
                    //    {
                    //        if (parameter["recipient"].Length > 0)
                    //        {
                    //            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "INTEGRA:NEWEMAIL:" + parameter["recipient"] + "|{00000000-0000-0000-0000-000000000000}|" + parameter["subject"], true);
                    //        }
                    //        else
                    //        {
                    //            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "INTEGRA:NEWEMAIL:|{00000000-0000-0000-0000-000000000000}|" + parameter["subject"], true);
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (parameter["recipient"].Length > 0)
                    //        {
                    //            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "INTEGRA:NEWEMAIL:" + parameter["recipient"] + "|{00000000-0000-0000-0000-000000000000}|", true);
                    //        }
                    //        else
                    //            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "INTEGRA:NEWEMAIL:|{00000000-0000-0000-0000-000000000000}|", true);
                    //    }
                    //}

                    if (parameter["action"] == "show")
                    {

                        if (parameter["guid"].Length == 36)
                        {
                            if (!Tools.CheckForAMTANGEE())
                                return;


                            AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "*AMTANGEE.Modules.Messages.dll:OPEN:{" + parameter["guid"] + "}");

                            //AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "INTEGRA:SHOWMESSAGE:{" + parameter["guid"] + "}", true);
                        }
                    }


                    if (parameter["action"] == "new")
                    {

                        if (!Tools.CheckForAMTANGEE())
                            return;

                        AMTANGEE.SDK.Messages.Email email = new SDK.Messages.Email(AMTANGEE.SDK.Global.CurrentUser.SpecialMessageCategories.Drafts);


                        if (parameter["subject"].Length > 0)
                            email.Subject = parameter["subject"];
                        if (parameter["recipient"].Length > 0)
                            email.To.Add(parameter["recipient"]);

                        //if (parameter["catrecipient"].Length > 0)
                        //    email.To.Add(parameter["catrecipient"]);

                        if (parameter["ccrecipient"].Length > 0)
                            email.To.Add( SDK.Messages.Email.Recipient.Modes.Cc, parameter["ccrecipient"]);

                        if (parameter["bccrecipient"].Length > 0)
                            email.To.Add(SDK.Messages.Email.Recipient.Modes.Bcc, parameter["bccrecipient"]);

                        if (parameter["plaintext"].Length > 0)
                            email.PlainText = parameter["plaintext"];
                        if (parameter["htmltext"].Length > 0)
                            email.HtmlText = parameter["htmltext"];


                        if (parameter["attachments"].Length > 0)
                        {
                            string att = parameter["attachments"];
                            if (att.Contains(";"))
                            {
                                System.Collections.Generic.List<string> atts = new List<string>();
                                string temp = att;

                                while (temp.Contains(";"))
                                {
                                    atts.Add(AMTANGEE.SDK.Global.CopyToPattern(temp, ";"));
                                    temp = AMTANGEE.SDK.Global.CopyFromPattern(temp, ";");
                                }
                                atts.Add(temp);

                                try
                                {
                                    foreach (string str in atts)
                                    {
                                        if (System.IO.File.Exists(str))
                                        {
                                            AMTANGEE.SDK.Messages.Attachment attachment = new AMTANGEE.SDK.Messages.Attachment(email);
                                            attachment.LoadFromFile(str);
                                            attachment.FileName = System.IO.Path.GetFileName(str);
                                            attachment.Save();
                                        }
                                    }
                                }
                                catch
                                {
                                }

                            }
                            else
                            {
                                try
                                {
                                    if (System.IO.File.Exists(att))
                                    {
                                        AMTANGEE.SDK.Messages.Attachment attachment = new AMTANGEE.SDK.Messages.Attachment(email);
                                        attachment.LoadFromFile(att);
                                        attachment.FileName = System.IO.Path.GetFileName(att);
                                        attachment.Save();
                                    }
                                }
                                catch
                                {
                                }
                            }
                        }

                        if (parameter["sender"].Trim().Length > 0)
                        {
                            AMTANGEE.SDK.Email.Accounts accounts = new SDK.Email.Accounts(AMTANGEE.SDK.Global.CurrentUser);
                            foreach (AMTANGEE.SDK.Email.Account acc in accounts)
                            {
                                if (acc.Send && acc.EmailAddress.Trim().ToUpper() == parameter["sender"].Trim().ToUpper())
                                {
                                    email.EmailAccount = acc;
                                    break;
                                }
                            }
                        }


                        if (parameter["template"].Trim().Length > 0)
                        {
                            AMTANGEE.SDK.Email.Templates templates = new SDK.Email.Templates(AMTANGEE.SDK.Global.CurrentUser);
                            foreach (AMTANGEE.SDK.Email.Template t in templates)
                            {
                                if (t.Name.ToUpper().Trim() == parameter["template"].Trim().ToUpper())
                                {
                                    email.Subject = t.Subject;
                                    email.PlainText = t.PlainText;
                                    email.HtmlText = t.HtmlTextToDisplay;
                                    
                                    
                                    using (AMTANGEE.SDK.TempDirectory tempDirectory = new AMTANGEE.SDK.TempDirectory())
                                    {
                                        foreach (AMTANGEE.SDK.Email.TemplateAttachment attachment in t.Attachments.AttachmentsWithoutContentId)
                                        {
                                            //editObject.Attachments.Add(attachment.SaveToTempFile(tempDirectory), attachment.ContentId, attachment.DeleteAfterSending).FileName = attachment.FileName;
                                            attachment.SaveToFile(tempDirectory.ToString() + attachment.Guid); email.Attachments.Add(tempDirectory.ToString() + attachment.Guid, attachment.ContentId, attachment.DeleteAfterSending).FileName = attachment.FileName; // {*}{+} sonst gibts z.T. Probleme mit zu langen Dateinamen, etc. !?! :-/
                                        }
                                    }

                                    email.SetSignature(null);
                                    break;
                                }
                            }
                        }
                        if (parameter["subject"].Length > 0)
                            email.Subject = parameter["subject"];
                        if (parameter["recipient"].Length > 0)
                            email.To.Add(parameter["recipient"]);

                        if (parameter["plaintext"].Length > 0)
                            email.PlainText = parameter["plaintext"];
                        if (parameter["htmltext"].Length > 0)
                            email.HtmlText = parameter["htmltext"];

                        AMTANGEE.SDK.Events.Send(new AMTANGEE.SDK.Users.User(new Guid(parameter["[USERGUID]"])), "*AMTANGEE.Modules.Messages.dll:OPEN:{" + email.Guid.ToString() + "}");
                        

                    }

                    if (parameter["action"] == "send")
                    {
                        AMTANGEE.SDK.Messages.Email email = new AMTANGEE.SDK.Messages.Email(AMTANGEE.SDK.Global.CurrentUser.SpecialMessageCategories.Outbound);


                        if (parameter["subject"].Length > 0)
                            email.Subject = parameter["subject"];
                        if (parameter["recipient"].Length > 0)
                            email.To.Add(parameter["recipient"]);
                        else
                            return;

                        if (parameter["plaintext"].Length > 0)
                            email.PlainText = parameter["plaintext"];
                        if (parameter["htmltext"].Length > 0)
                            email.HtmlText = parameter["htmltext"];

                        if (parameter["attachments"].Length > 0)
                        {
                            string att = parameter["attachments"];
                            if (att.Contains(";"))
                            {
                                System.Collections.Generic.List<string> atts = new List<string>();
                                string temp = att;

                                while (temp.Contains(";"))
                                {
                                    atts.Add(AMTANGEE.SDK.Global.CopyToPattern(temp, ";"));
                                    temp = AMTANGEE.SDK.Global.CopyFromPattern(temp, ";");
                                }
                                atts.Add(temp);

                                try
                                {
                                    foreach (string str in atts)
                                    {
                                        if (System.IO.File.Exists(str))
                                        {
                                            AMTANGEE.SDK.Messages.Attachment attachment = new AMTANGEE.SDK.Messages.Attachment(email);
                                            attachment.LoadFromFile(str);
                                            attachment.FileName = System.IO.Path.GetFileName(str);
                                            attachment.Save();
                                        }
                                    }
                                }
                                catch
                                {
                                }

                            }
                            else
                            {
                                try
                                {
                                    if (System.IO.File.Exists(att))
                                    {
                                        AMTANGEE.SDK.Messages.Attachment attachment = new AMTANGEE.SDK.Messages.Attachment(email);
                                        attachment.LoadFromFile(att);
                                        attachment.FileName = System.IO.Path.GetFileName(att);
                                        attachment.Save();
                                    }
                                }
                                catch
                                {
                                }
                            }
                        }

                        if (parameter["approved"].Length > 0)
                        {
                            if (parameter["approved"].ToLower() == "true")
                                email.Approved = true;

                            if (parameter["approved"].ToLower() == "false")
                                email.Approved = false;
                        }
                        else
                            email.Approved = false;


                        if (parameter["signature"].Length > 0)
                        {


                            if (parameter["signature"].ToLower() == "false")
                                email.SetSignature(null);
                            else
                            {
                                if (email.EmailAccount.DefaultSignature != null && email.EmailAccount.DefaultSignature.ExistsAndLoadedAndRights)
                                {
                                    email.SetSignature(email.EmailAccount.DefaultSignature);
                                }
                            }

                        }
                        else
                        {
                            if (email.EmailAccount.DefaultSignature != null && email.EmailAccount.DefaultSignature.ExistsAndLoadedAndRights)
                            {
                                email.SetSignature(email.EmailAccount.DefaultSignature);
                            }
                        }
                     
                           

                        bool multiMail = false;
                        if (parameter["mailtype"].Length > 0)
                        {
                            if (parameter["mailtype"].ToLower() == "single")
                                multiMail = false;

                            if (parameter["mailtype"].ToLower() == "multi")
                                multiMail = true;
                        }
                        //try
                        //{
                     
                        email.Send(multiMail);
                        //}
                        //catch
                        //{
                        //}
                    }
                }

                public string ModuleName
                {
                    get { return "EMAIL"; }
                }
            }


            #endregion
        }
    }
}