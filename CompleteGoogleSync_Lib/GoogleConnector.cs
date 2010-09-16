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
using System.Collections;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Extensions;
using Google.GData.Client;
using Google.Contacts;

namespace GoogleSynchronizer
{
    public partial class GoogleConnector
    {
        private string _username;
        private string _password;
        private Hashtable _contactList = new Hashtable();
        private Hashtable _calendarItemsList = new Hashtable();
        private Hashtable _contactsGroupsList = new Hashtable();

        private RequestSettings requestSettings;
        private ContactsRequest contactsRequest;

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

        public string Username
        {
            get { return _username; }
            set { _username = value; }
        }

        public string Password
        {
            get { return _password; }
            set { _password = value; }
        }

        public void InitializeGoogleAccountService()
        {
            requestSettings = new RequestSettings("", Username, Password);
            contactsRequest = new ContactsRequest(requestSettings);
        }
    }
}
