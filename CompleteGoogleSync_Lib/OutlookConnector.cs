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
using System.Reflection;
using System.Collections;

namespace GoogleSynchronizer
{
    public partial class OutlookConnector
    {
        private Application olApp;
        private NameSpace olNameSpace;

        private Hashtable _contactList = new Hashtable();
        private Hashtable _calendarItemsList = new Hashtable();
        private Hashtable _contactsGroupsList = new Hashtable();

        private bool _useCustomCategory;

        public bool UseCustomCatergory
        {
            get { return _useCustomCategory; }
            set { _useCustomCategory = value; }
        }

        public Hashtable ContactsGroupsList
        {
            get { return _contactsGroupsList; }
        }

        public Hashtable ContactList
        {
            get { return _contactList; }
        }

        public Hashtable CalendarItemList
        {
            get { return _calendarItemsList; }
        }

        public OutlookConnector()
        {
            olApp = new Outlook.Application();
            olNameSpace = olApp.GetNamespace("MAPI");
            ContactItem contact = (ContactItem)olApp.CreateItem(OlItemType.olContactItem);
            _outlookNullDateValue = contact.Birthday;
            contact = null;

            if (!VerifyOutlookVersion())
            {
                throw new ArgumentException("Incorrect Outlook version");
            }
        }

        public bool VerifyOutlookVersion()
        {
            if (olApp.Version.StartsWith("11"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        ~OutlookConnector()
        {
            olNameSpace = null;
            olApp = null;
        }
    }
}
