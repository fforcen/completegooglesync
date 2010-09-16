/*
    Copyright (C) 2010 by Fernando Forcén López

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.Text;
using Outlook;
using Google.Contacts;
using System.Collections;
using System.IO;
using Encriptacion;

namespace GoogleSynchronizer
{    /// <summary>
    /// Priority enumeration in conflict case
    /// </summary>
    public enum SyncPriority
    {
        Google =0,
        Outlook=1,
        Preguntar=2
    }

    /// <summary>
    /// Synchronization status enumeration
    /// </summary>
    public enum SynchronizationStatus{
        Synchronized,
        SynchronizingContacts,
        SynchronizingCalendar,
        SynchronizationError
    }

    /// <summary>
    /// Synchronizer class
    /// </summary>
    public partial class Synchronizer
    {
        private Encriptador encrypting = new Encriptador();
        private string _username;
        private string _password;
        private OutlookConnector outlookConnector;
        private GoogleConnector googleConnector;
        private SyncPriority _syncPriority;

        private string synchronizationFile = System.IO.Path.GetFullPath(".") + "\\" + SynchronizerConfig.CONFIGURATION_FILE;
        private string logFile = System.IO.Path.GetFullPath(".") + "\\" + SynchronizerConfig.LOG_FILE;

        private StreamWriter logWriter;

        SyncStatus syncStatus = new SyncStatus();

        public delegate void ChangeStatusEventHandler(SynchronizationStatus newStatus, String message);

        public event ChangeStatusEventHandler ChangeStatusEvent;

        public String StoreGoogleAccount
        {
            get { return syncStatus.GoogleAccount; }
        }

        public String StoreGoogleAccountPassword
        {
            get
            {
                if (syncStatus.GoogleAccountPassword != null)
                {
                    return encrypting.DecryptString128Bit(syncStatus.GoogleAccountPassword,SynchronizerConfig.ENCRYPTION_PASSWORD);
                }
                else
                {
                    return "";
                }
            }
        }

        public bool UseCustomCatergory
        {
            get { return syncStatus.UseCustomCategory; }
            set { syncStatus.UseCustomCategory = value; }
        }

        public SyncPriority SyncPriority
        {
            get { return _syncPriority; }
            set { _syncPriority = value; }
        }

        public string GoogleAccountUsername
        {
            get { return _username; }
            set { _username = value; syncStatus.GoogleAccount = _username; }
        }

        public string GoogleAccountPassword
        {

            get { return _password; }
            set { _password = value; syncStatus.GoogleAccountPassword = encrypting.EncryptString128Bit(_password,SynchronizerConfig.ENCRYPTION_PASSWORD); }
        }

        /// <summary>
        /// Main method to synchronize Google Account and Outlook
        /// </summary>
        public void Synchronize()
        {
            try
            {
                googleConnector = new GoogleConnector();
                outlookConnector = new OutlookConnector();

                outlookConnector.UseCustomCatergory = this.UseCustomCatergory;
                googleConnector.UseCustomCatergory = this.UseCustomCatergory;

                googleConnector.Username = _username;
                googleConnector.Password = _password;

                WriteLog("Inciando sincronización");

                ChangeStatusEvent(SynchronizationStatus.SynchronizingContacts, null);
                WriteLog("Iniciando sincronizacion de contactos");

                googleConnector.InitializeGoogleAccountService();
                SynchronizeContacts();

                WriteLog("Fin sincronizacion de contactos");

                ChangeStatusEvent(SynchronizationStatus.Synchronized, null);
                SaveSyncStatus();
                WriteLog("Sincronizacion finalizada");

                outlookConnector = null;
                googleConnector = null;
            }
            catch (System.Exception e)
            {
                WriteLog("ERROR: " + e.Message);
                ChangeStatusEvent(SynchronizationStatus.SynchronizationError, e.Message);
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public Synchronizer()
        {
            logWriter = File.AppendText(logFile);
            LoadSyncStatus();
        }

        /// <summary>
        /// Synchronize contacts intems
        /// </summary>
        private void SynchronizeContacts()
        {
            try
            {
                googleConnector.OutlookNullDateValue = outlookConnector.OutlookNullDateValue;

                outlookConnector.LoadContacts();
                WriteLog("Outlook: Se han recuperado " + outlookConnector.ContactList.Count + " contactos");

                googleConnector.LoadContactsGroups();

                googleConnector.LoadContacts();
                WriteLog("Google: Se han recuperado " + googleConnector.ContactList.Count + " contactos");

                foreach (Contact currentContact in googleConnector.ContactList.Values)
                {
                    CheckOutlookContact(currentContact);
                }

                foreach (ContactItem currentContact in outlookConnector.ContactList.Values)
                {
                    CheckGoogleContact(currentContact);
                }
            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Check a Google Contact against the Outlook contacts list
        /// </summary>
        /// <param name="gContact">Google contact to check</param>
        private void CheckOutlookContact(Contact gContact)
        {
            ContactItem oContact = null;
            RelationItem basicContact = null;

            String id;
            String relatedId;

            DateTime oContactDate;
            DateTime gContactDate;
            DateTime sDate;

            id = gContact.Id;

            // Check if exist in syncrhonization list
            if (syncStatus.contactList.ContainsKey(gContact.Id)){
                // Yes: Find reltacion object
                basicContact = (RelationItem)syncStatus.contactList[gContact.Id];
                if (outlookConnector.ContactList.ContainsKey(basicContact.RelatedId)) {
                    oContact = (ContactItem) outlookConnector.ContactList[basicContact.RelatedId];
                }
                else{
                    //No: remove contact
                    syncStatus.contactList.Remove(gContact.Id);
                    googleConnector.RemoveContact(gContact);
                    return;
                }
            }
            else {
                // Check if exist by name
                foreach (ContactItem contact in outlookConnector.ContactList.Values)
                {
                    if (contact.FullName!=null && contact.FullName.Equals(gContact.Name.FullName))
                    {
                        oContact = contact;
                        break;
                    }
                }
            }

            if (oContact == null)
            {
                // No exist in synchronization list, therefore is new
                relatedId = outlookConnector.SaveContactFromGoogleContact(gContact,googleConnector.ContactsGroupsList);

                ContactListUpdate(id, ContactSource.Google,relatedId);
                ContactListUpdate(relatedId, ContactSource.Outlook, id);
            }
            else
            {
                // Exist, check synchronization date/time
                oContactDate = oContact.LastModificationTime;
                gContactDate = gContact.Updated;
                sDate = syncStatus.SynchronizationTime;

                if (DateTime.Compare(gContactDate,oContactDate) < 0 &&  DateTime.Compare(oContactDate,sDate) < 0)
                {
                    // Synchronize!!!, nothing to do
                }
                else if (DateTime.Compare(gContactDate,sDate) > 0 &&  DateTime.Compare(sDate,oContactDate)>0)
                {
                    // Google Contact latest than outlook contact
                    // Update outlook contact
                    outlookConnector.UpdateContactFromGoogleContact(gContact, oContact, googleConnector.ContactsGroupsList);
                    relatedId = oContact.EntryID;

                    ContactListUpdate(id, ContactSource.Google, relatedId);
                    ContactListUpdate(relatedId, ContactSource.Outlook, id);
                }
                else if (DateTime.Compare(oContactDate, sDate) > 0 && DateTime.Compare(sDate, gContactDate) > 0)
                {
                    // El contacto de outlook es más reciente que el de google
                    // no hacemos nada porque se actualizara cuando se recorra
                    // la lista de outlook
                }
                else if (DateTime.Compare(gContactDate, oContactDate) > 0 && DateTime.Compare(oContactDate, sDate) > 0)
                {
                    // Google Contact and Outlook contact both outdate
                    // CASO: AMBOS ACTUALIZADOS
                }
            }
        }

        /// <summary>
        /// Check an Outlook Contact against the Google contacts list
        /// </summary>
        /// <param name="oContact"></param>
        private void CheckGoogleContact(ContactItem oContact)
        {
            Contact gContact = null;
            RelationItem basicContact = null;

            String id;
            String relatedId;

            DateTime oContactDate;
            DateTime gContactDate;
            DateTime sDate;

            id = oContact.EntryID;

            // Check if exist in syncrhonization list
            if (syncStatus.contactList.ContainsKey(oContact.EntryID))
            {
                // Yes: Find reltacion object
                basicContact = (RelationItem)syncStatus.contactList[oContact.EntryID];
                if (googleConnector.ContactList.ContainsKey(basicContact.RelatedId))
                {
                    gContact = (Contact)googleConnector.ContactList[basicContact.RelatedId];
                }
                else
                {
                    //No: remove contact
                    syncStatus.contactList.Remove(oContact.EntryID);
                    outlookConnector.RemoveContact(oContact);
                    return;
                }
            }
            else
            {
                // Check if exist by name
                foreach (Contact contact in googleConnector.ContactList.Values)
                {
                    if (contact.Name.FullName!=null && contact.Name.FullName.Equals(oContact.FullName))
                    {
                        gContact = contact;
                        break;
                    }
                }
            }

            if (gContact == null)
            {
                // No exist in synchronization list, therefore is new
                relatedId = googleConnector.SaveContactFromOutlookContact(oContact);

                ContactListUpdate(id, ContactSource.Outlook, relatedId);
                ContactListUpdate(relatedId, ContactSource.Google, id);
            }
            else
            {
                // Exist, check synchronization date/time
                oContactDate = oContact.LastModificationTime;
                gContactDate = gContact.Updated;
                sDate = syncStatus.SynchronizationTime;

                if (DateTime.Compare(gContactDate, oContactDate) < 0 && DateTime.Compare(oContactDate, sDate) < 0)
                {
                    // Synchronize!!!, nothing to do
                }
                else if (DateTime.Compare(oContactDate, sDate) > 0 && DateTime.Compare(sDate, gContactDate) > 0)
                {
                    // Outlook Contact latest than google contact
                    // Update google contact
                    googleConnector.UpdateContactFromOutlookContact(gContact, oContact);
                    relatedId = gContact.Id;

                    ContactListUpdate(id, ContactSource.Outlook, relatedId);
                    ContactListUpdate(relatedId, ContactSource.Google, id);
                }
                else if (DateTime.Compare(gContactDate, sDate) > 0 && DateTime.Compare(sDate, oContactDate) > 0)
                {
                    // El contacto de outlook es más reciente que el de google
                    // no hacemos nada porque se actualizara cuando se recorra
                    // la lista de outlook
                }
                else if (DateTime.Compare(gContactDate, oContactDate) > 0 && DateTime.Compare(oContactDate, sDate) > 0)
                {
                    // Ambos contactos se han modificado
                    // CASO: AMBOS ACTUALIZADOS
                }
            }
        }

        /// <summary>
        /// Update the synchronization list
        /// </summary>
        /// <param name="id">Id of the synchronized object</param>
        /// <param name="source">Source of the item</param>
        /// <param name="relatedid">Related item of this item</param>
        public void ContactListUpdate(String id, ContactSource source, String relatedid)
        {
            RelationItem basicContact;

            if (!syncStatus.contactList.ContainsKey(id))
            {
                basicContact = new RelationItem();
                basicContact.Id = id;
                basicContact.RelatedId = relatedid;
                basicContact.Source = source;
                syncStatus.contactList.Add(id, basicContact);
            }
        }

        

    }

}
