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
using System.IO;
using System.Collections;
using System.Xml;
using System.Xml.Serialization;

namespace GoogleSynchronizer
{

    /// <summary>
    /// Enumeration of synchronization status of each item
    /// </summary>
    public enum BasicContactStatus{
        Syncronized,
        OutofDate,
        New
    }

    /// <summary>
    /// Source of the contact item
    /// </summary>
    public enum ContactSource
    {
        Google,
        Outlook
    }

    /// <summary>
    /// Class to manage configuration XML
    /// </summary>
    [XmlRoot("Configuration")]
    public class SyncStatus
    {
        protected internal Hashtable contactList;
        protected internal Hashtable calendarItemList;

        [XmlAttribute("SynchronizationTime")]
        private DateTime synchronizationTime;
        [XmlAttribute("GoogleAccount")]
        private String googleAccount;
        [XmlAttribute("GoogleAccountPassword")]
        private String googleAccountPassword;
        [XmlAttribute("UseCustomCategory")]
        private Boolean useCustomCategory;

        public SyncStatus()
        {
            contactList = new Hashtable();
            calendarItemList = new Hashtable();
            synchronizationTime = DateTime.MinValue;
        }

        public Boolean UseCustomCategory
        {
            get { return useCustomCategory; }
            set { useCustomCategory = value; }
        }

        public String GoogleAccount
        {
            get { return googleAccount; }
            set { googleAccount = value; }
        }

        public String GoogleAccountPassword
        {
            get { return googleAccountPassword; }
            set { googleAccountPassword = value; }
        }

        public DateTime SynchronizationTime
        {
            get { return synchronizationTime; }
            set { synchronizationTime = value; }
        }

        [XmlElement("contact")]
        public RelationItem[] Contacts
        {
            get
            {
                RelationItem[] contacts = new RelationItem[contactList.Values.Count];
                contactList.Values.CopyTo(contacts, 0);
                return contacts;
            }
            set
            {
                if (value == null) return;
                RelationItem[] contacts = (RelationItem[])value;
                contactList.Clear();
                foreach (RelationItem contact in contacts)
                    contactList.Add(contact.Id, contact);
            }
        }

        [XmlElement("calendaritem")]
        public RelationItem[] CalendarItems
        {
            get
            {
                RelationItem[] calendarItems = new RelationItem[calendarItemList.Values.Count];
                calendarItemList.Values.CopyTo(calendarItems, 0);
                return calendarItems;
            }
            set
            {
                if (value == null) return;
                RelationItem[] calendarItems = (RelationItem[])value;
                calendarItemList.Clear();
                foreach (RelationItem calendarItem in calendarItems)
                    calendarItemList.Add(calendarItem.Id, calendarItem);
            }
        }
    }

    /// <summary>
    /// Class that represents a relation between two contacts
    /// </summary>
    public class RelationItem
    {
        [XmlAttribute("id")]
        private String _id;
        [XmlAttribute("relatedid")] 
        private String _relatedId;
        [XmlAttribute("source")]
        private ContactSource _source;

        public String Id{
            get { return _id; }
            set { _id = value; }
        }

        public ContactSource Source
        {
            get { return _source; }
            set { _source = value; }
        }

        public String RelatedId
        {
            get { return _relatedId; }
            set { _relatedId = value; }
        }
    }
    
    public partial class Synchronizer
    {
        /// <summary>
        /// Load configuration file
        /// </summary>
        private void LoadSyncStatus()
        {
            try
            {
                XmlSerializer s = new XmlSerializer(typeof(SyncStatus));
                TextReader r = new StreamReader(synchronizationFile);
                syncStatus = (SyncStatus)s.Deserialize(r);
                r.Close();
            }
            catch (Exception)
            {
                WriteLog("ERROR: No se ha podido recuperar el registro de sincronizacion");
            }
        }

        /// <summary>
        /// Save configuration file
        /// </summary>
        private void SaveSyncStatus()
        {
            try
            {
                syncStatus.SynchronizationTime = DateTime.Now;

                XmlSerializer s = new XmlSerializer(typeof(SyncStatus));
                TextWriter w = new StreamWriter(synchronizationFile);
                s.Serialize(w, syncStatus);
                w.Close();
            }
            catch (Exception)
            {
                WriteLog("ERROR: No se ha podido actualizar el registro de sincronizacion");
            }
        }

        /// <summary>
        /// Write a line into log file
        /// </summary>
        /// <param name="message">Message to write</param>
        private void WriteLog(String message)
        {
            logWriter.WriteLine("{0} {1} {2}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString(), message);
            logWriter.Flush();
        }
    }
}
