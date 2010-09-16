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
using Google.GData.Extensions;
using System.Reflection;
using System.Collections;

namespace GoogleSynchronizer
{
    public partial class OutlookConnector
    {
        private DateTime _outlookNullDateValue;
        private string googleToken;

        public string GoogleToken
        {
            get { return googleToken; }
            set { googleToken = value; }
        }

        public DateTime OutlookNullDateValue
        {
            get { return _outlookNullDateValue; }
            set { _outlookNullDateValue = value; }
        }

        /// <summary>
        /// Remove outlook contact
        /// </summary>
        /// <param name="oContact">Outlook contacto object to remove</param>
        public void RemoveContact(ContactItem oContact)
        {
            oContact.Delete();
        }

        /// <summary>
        /// Create and save a Outlook contact from a Goolge Contact Object
        /// </summary>
        /// <param name="gContact">Google Contact object</param>
        /// <param name="googleGroups">List of all Google Groups</param>
        /// <returns>Id from created Outlook contact</returns>
        public string SaveContactFromGoogleContact(Contact gContact, Hashtable googleGroups)
        {
            string id;
            ContactItem oContact;
            oContact = (ContactItem)olApp.CreateItem(OlItemType.olContactItem);
            GoogleContactToOutlookContact(oContact, gContact, googleGroups);
            oContact.Save();
            id = oContact.EntryID;
            oContact = null;
            return id;
        }

        /// <summary>
        /// Update Outlook Contact object data from Google Contact Object
        /// </summary>
        /// <param name="gContact">Google Object</param>
        /// <param name="oContact">Outlook Object</param>
        /// <param name="googleGroups">List of all Google Groups</param>
        public void UpdateContactFromGoogleContact(Contact gContact, ContactItem oContact, Hashtable googleGroups)
        {
            GoogleContactToOutlookContact(oContact, gContact, googleGroups);
            oContact.Save();
            oContact = null;
        }

        /// <summary>
        /// Asign Outlook Contact data to Google Contact Object.
        /// </summary>
        /// <param name="oContact">Outlook contact object</param>
        /// <param name="gContact">Google contact object</param>
        /// <param name="googleGroups">List of all Google Groups</param>
        private void GoogleContactToOutlookContact(ContactItem oContact, Contact gContact, Hashtable googleGroups)
        {
            int counter;
            string categories;

            // Categories
            categories="";
            foreach (Google.GData.Contacts.GroupMembership msGroup in gContact.GroupMembership)
            {
                Group group = (Group) googleGroups[msGroup.HRef];
                /*if (group.SystemGroup==null)
                {
                    categories = categories + group.Title;
                }
                else
                {
                    categories = categories + group.SystemGroup;
                }*/
                categories = categories + group.Title;
                categories = categories + "; ";
            }
            if (UseCustomCatergory)
            {
                categories = categories + SynchronizerConfig.CUSTOM_CATEGORY + "; ";
            }
            if (categories.Length > 0)
            {
                categories = categories.Substring(0, categories.Length - 2);
                oContact.Categories = categories;
            }


            //Telephones
            foreach (Google.GData.Extensions.PhoneNumber number in gContact.Phonenumbers)
            {
                if (number.Primary)
                {
                    oContact.PrimaryTelephoneNumber = number.Value.ToString();
                }
                
                if (number.Home)
                {
                    oContact.HomeTelephoneNumber = number.Value.ToString();
                }
                else if (number.Work)
                {
                    oContact.BusinessTelephoneNumber = number.Value.ToString();
                }
                else
                {
                    switch (number.Rel)
                    {
                        case ContactsRelationships.IsMobile:
                            {
                                oContact.MobileTelephoneNumber = number.Value.ToString();
                                break;
                            }
                        case ContactsRelationships.IsHomeFax:
                            {
                                oContact.HomeFaxNumber = number.Value.ToString();
                                break;
                            }
                        case ContactsRelationships.IsHome:
                            {
                                oContact.HomeTelephoneNumber = number.Value.ToString();
                                break;
                            }
                        case ContactsRelationships.IsFax:
                            {
                                oContact.BusinessFaxNumber = number.Value.ToString();
                                break;
                            }
                        case ContactsRelationships.IsWork:
                            {
                                oContact.BusinessTelephoneNumber = number.Value.ToString();
                                break;
                            }
                        case ContactsRelationships.IsOther:
                            {
                                oContact.OtherTelephoneNumber = number.Value.ToString();
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
            }

            //Address
            foreach (Google.GData.Extensions.StructuredPostalAddress address in gContact.PostalAddresses)
            {
                switch (address.Rel)
                {
                    case ContactsRelationships.IsHome:
                        {
                            if (address.FormattedAddress != null) { oContact.HomeAddress = address.FormattedAddress; }
                            if (address.City != null) { oContact.HomeAddressCity = address.City; }
                            if (address.Country != null) { oContact.HomeAddressCountry = address.Country; }
                            if (address.Postcode != null) { oContact.HomeAddressPostalCode = address.Postcode; }
                            if (address.Pobox != null) { oContact.HomeAddressPostOfficeBox = address.Pobox; }
                            if (address.Subregion != null) { oContact.HomeAddressState = address.Subregion; }
                            if (address.Street!=null) { oContact.HomeAddressStreet = address.Street; }
                            break;
                        }
                    case ContactsRelationships.IsWork:
                        {
                            if (address.FormattedAddress != null) { oContact.BusinessAddress = address.FormattedAddress; }
                            if (address.City != null) { oContact.BusinessAddressCity = address.City; }
                            if (address.Country != null) { oContact.BusinessAddressCountry = address.Country; }
                            if (address.Postcode != null) { oContact.BusinessAddressPostalCode = address.Postcode; }
                            if (address.Pobox != null) { oContact.BusinessAddressPostOfficeBox = address.Pobox; }
                            if (address.Subregion != null) { oContact.BusinessAddressState = address.Subregion; }
                            if (address.Street != null) { oContact.BusinessAddressStreet = address.Street; }
                            break;
                        }
                    case ContactsRelationships.IsOther:
                        {
                            if (address.FormattedAddress != null) { oContact.OtherAddress = address.FormattedAddress; }
                            if (address.City != null) { oContact.OtherAddressCity = address.City; }
                            if (address.Country != null) { oContact.OtherAddressCountry = address.Country; }
                            if (address.Postcode != null) { oContact.OtherAddressPostalCode = address.Postcode; }
                            if (address.Pobox != null) { oContact.OtherAddressPostOfficeBox = address.Pobox; }
                            if (address.Subregion != null) { oContact.OtherAddressState = address.Subregion; }
                            if (address.Street != null) { oContact.OtherAddressStreet = address.Street; }
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }

            oContact.FullName = gContact.Name.FullName != null ? gContact.Name.FullName.ToString() : null;
            oContact.Account = gContact.PrimaryEmail != null ? gContact.PrimaryEmail.Address.ToString() : null;
            oContact.Birthday  = gContact.ContactEntry.Birthday != null ? DateTime.Parse(gContact.ContactEntry.Birthday) : _outlookNullDateValue;
            oContact.Body = gContact.Content != null ? gContact.Content.ToString(): null;
            if (gContact.Organizations.Count > 0)
            {
                oContact.CompanyName = gContact.Organizations[0].Name != null ? gContact.Organizations[0].Name.ToString(): null;
                oContact.JobTitle = gContact.Organizations[0].Title != null ? gContact.Organizations[0].Title.ToString() : null;
            }
            else
            {
                oContact.CompanyName = null;
                oContact.JobTitle = null;
            }
            //Emails
            counter = 1;
            foreach (EMail email in gContact.Emails)
            {
                switch (counter)
                {
                    case 1:
                        oContact.Email1Address = email.Address.ToString();
                        break;
                    case 2:
                        oContact.Email2Address = email.Address.ToString();
                        break;
                    case 3:
                        oContact.Email3Address = email.Address.ToString();
                        break;
                    default:
                        counter=gContact.Emails.Count;
                        break;
                }
                counter++;
            }
            oContact.IMAddress = gContact.IMs.Count > 0 ? gContact.IMs[0].Address.ToString() : null;
            oContact.Initials  = gContact.ContactEntry.Initials != null ? gContact.ContactEntry.Initials.ToString() : null;
            oContact.NickName  = gContact.ContactEntry.Nickname != null ? gContact.ContactEntry.Nickname.ToString() : null;
            oContact.WebPage  = gContact.ContactEntry.Websites.Count > 0 ? gContact.ContactEntry.Websites[0].Href.ToString() : null;
        }

        /// <summary>
        /// Load all Outlook contacts into a list
        /// </summary>
        public void LoadContacts()
        {
            _contactList = new Hashtable();

            MAPIFolder cContacts = olNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

            foreach (ContactItem contact in cContacts.Items)
            {
                _contactList.Add(contact.EntryID, contact);
            }
        }
    }
}
