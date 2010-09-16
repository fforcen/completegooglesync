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
using System.IO;
using System.Net;
using Google.GData.Contacts;
using Google.GData.Extensions;
using Google.GData.Client;
using Google.Contacts;
using Outlook;
using System.Collections.Generic;

namespace GoogleSynchronizer
{
    public partial class GoogleConnector
    {
        private DateTime _outlookNullDateValue;

        /// <summary>
        /// Default date to null value date in outlook contact
        /// </summary>
        public DateTime OutlookNullDateValue
        {
            get { return _outlookNullDateValue; }
            set { _outlookNullDateValue = value; }
        }

        /// <summary>
        /// Load contacts groups in list
        /// </summary>
        public void LoadContactsGroups()
        {
            _contactsGroupsList = new Hashtable();

            requestSettings.AutoPaging = true;

            Feed<Group> fg = contactsRequest.GetGroups();
            try
            {
                foreach (Group g in fg.Entries)
                {
                    _contactsGroupsList.Add(g.Id, g);
                }
            }
            catch (GDataRequestException e)
            {
                throw new ArgumentException("Error procesando la solicitud");
            }
            catch (System.Exception e)
            {
                
                throw e;
            }
        }

        /// <summary>
        /// Load all google contacts into a list
        /// </summary>
        public void LoadContacts()
        {
            _contactList = new Hashtable();

            // AutoPaging results in automatic paging in order to retrieve all contacts
            requestSettings.AutoPaging = true; 

            Feed<Contact> f = contactsRequest.GetContacts();

            try
            {
                foreach (Contact e in f.Entries)
                {
                    _contactList.Add(e.Id,e);
                }      
            }
            catch (GDataRequestException e)
            {
                throw e;
            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Update Google Contact object data from Outlook Contact Object
        /// </summary>
        /// <param name="gContact">Google Contact Object</param>
        /// <param name="oContact">Outlook Contact Object</param>
        public void UpdateContactFromOutlookContact(Contact gContact,ContactItem oContact)
        {
            OutlookContactToGoogleContact(oContact, gContact);
            contactsRequest.Update(gContact);
        }

        /// <summary>
        /// Remove Google Contact
        /// </summary>
        /// <param name="gContact">Google Contact Object to remove</param>
        public void RemoveContact(Contact gContact)
        {
            contactsRequest.Delete(gContact);
        }

        /// <summary>
        /// Create Google Group
        /// </summary>
        /// <param name="name">Group name</param>
        /// <returns>Created group</returns>
        public Group CreateGroup(String name)
        {
            
            Group newGroup = new Group();
            newGroup.Title = name;
            Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));
            Group createdGroup = contactsRequest.Insert(feedUri, newGroup);

            ContactsGroupsList.Add(createdGroup.Id, createdGroup);

            return createdGroup;
        }

        /// <summary>
        /// Create and save an Google Contact from Outlook Contact data
        /// </summary>
        /// <param name="oContact">Ooutlook contact object</param>
        /// <returns>Id from created Google Contact</returns>
        public string SaveContactFromOutlookContact(ContactItem oContact)
        {
            string id="";
            Contact gContact = new Contact();
            OutlookContactToGoogleContact(oContact, gContact);
            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
            Contact createdContact = contactsRequest.Insert(feedUri, gContact);
            id = createdContact.Id;
            
            return id;
        }

        /// <summary>
        /// Asign Outlook Contact data to Google Contact Object.
        /// </summary>
        /// <param name="oContact">Outlook object</param>
        /// <param name="gContact">Google object</param>
        private void OutlookContactToGoogleContact(ContactItem oContact, Contact gContact)
        {
            Website web;
            Organization organ;
            EMail email;
            PhoneNumber telephone;
            StructuredPostalAddress address;
            IMAddress imAddress;
            
            GroupMembership memberShip;
            Group currentGroup;
            List<GroupMembership> removedGroups = new List<GroupMembership>();

            //

            //Reset google contact
            gContact.Phonenumbers.Clear();
            gContact.PostalAddresses.Clear();
            gContact.ContactEntry.Websites.Clear();
            gContact.Emails.Clear();
            gContact.IMs.Clear();

            // Check categories to remove removed categories
            foreach (Google.GData.Contacts.GroupMembership msGroup in gContact.GroupMembership)
            {
                Group group = (Group)ContactsGroupsList[msGroup.HRef];
                if (!oContact.Categories.Contains(group.Title))
                {
                    removedGroups.Add(msGroup);
                }
            }
            // Remove removed Group
            foreach (GroupMembership ms in removedGroups)
            {
                gContact.GroupMembership.Remove(ms);
            }

            // Categories
            if (oContact.Categories != null)
            {
                string[] categories = oContact.Categories.Split(';');
                foreach (string category in categories)
                {
                    if (!category.Trim().Equals(SynchronizerConfig.CUSTOM_CATEGORY))
                    {
                        // Check in google contact
                        currentGroup = FindGroupInContact(gContact, category.Trim());
                        if (currentGroup == null)
                        {
                            memberShip = new GroupMembership();

                            //Check group in grouplist
                            currentGroup = FindGroup(category.Trim());
                            if (currentGroup == null)
                            {
                                currentGroup = CreateGroup(category.Trim());
                                memberShip.HRef = currentGroup.Id;
                            }
                            else
                            {
                                memberShip.HRef = currentGroup.Id;
                            }
                            gContact.GroupMembership.Add(memberShip);
                        }
                    }
                }
            }
            //Default Category
            if (UseCustomCatergory)
            {
                // Check in google contact
                currentGroup = FindGroupInContact(gContact, SynchronizerConfig.CUSTOM_CATEGORY);
                if (currentGroup == null)
                {
                    memberShip = new GroupMembership();

                    //Check group in grouplist
                    currentGroup = FindGroup(SynchronizerConfig.CUSTOM_CATEGORY);
                    if (currentGroup == null)
                    {
                        currentGroup = CreateGroup(SynchronizerConfig.CUSTOM_CATEGORY);
                        memberShip.HRef = currentGroup.Id;
                    }
                    else
                    {
                        memberShip.HRef = currentGroup.Id;
                    }
                    gContact.GroupMembership.Add(memberShip);
                }
            }

            //Telephones            
            if (oContact.HomeTelephoneNumber != null)
            {
                telephone = new PhoneNumber(oContact.HomeTelephoneNumber.ToString());
                telephone.Rel = ContactsRelationships.IsHome;
                if (oContact.HomeTelephoneNumber.Equals(oContact.PrimaryTelephoneNumber))
                {
                    telephone.Primary = true;
                }
                gContact.Phonenumbers.Add(telephone);
            }
            if (oContact.BusinessTelephoneNumber != null)
            {
                telephone = new PhoneNumber(oContact.BusinessTelephoneNumber.ToString());
                telephone.Rel = ContactsRelationships.IsWork;
                if (oContact.BusinessTelephoneNumber.Equals(oContact.PrimaryTelephoneNumber))
                {
                    telephone.Primary = true;
                }
                gContact.Phonenumbers.Add(telephone);
            }
            if (oContact.MobileTelephoneNumber != null)
            {
                telephone = new PhoneNumber(oContact.MobileTelephoneNumber.ToString());
                telephone.Rel = ContactsRelationships.IsMobile;
                if (oContact.MobileTelephoneNumber.Equals(oContact.PrimaryTelephoneNumber))
                {
                    telephone.Primary = true;
                }
                gContact.Phonenumbers.Add(telephone);
            }
            if (oContact.HomeFaxNumber != null)
            {
                telephone = new PhoneNumber(oContact.HomeFaxNumber.ToString());
                telephone.Rel = ContactsRelationships.IsHomeFax;
                if (oContact.HomeFaxNumber.Equals(oContact.PrimaryTelephoneNumber))
                {
                    telephone.Primary = true;
                }
                gContact.Phonenumbers.Add(telephone);
            }
            if (oContact.BusinessFaxNumber != null)
            {
                telephone = new PhoneNumber(oContact.BusinessFaxNumber.ToString());
                telephone.Rel = ContactsRelationships.IsWorkFax;
                if (oContact.BusinessFaxNumber.Equals(oContact.PrimaryTelephoneNumber))
                {
                    telephone.Primary = true;
                }
                gContact.Phonenumbers.Add(telephone);
            }
            if (oContact.OtherTelephoneNumber != null)
            {
                telephone = new PhoneNumber(oContact.OtherTelephoneNumber.ToString());
                telephone.Rel = ContactsRelationships.IsOther;
                if (oContact.OtherTelephoneNumber.Equals(oContact.PrimaryTelephoneNumber))
                {
                    telephone.Primary = true;
                }
                gContact.Phonenumbers.Add(telephone);
            }

            //Address
            if (oContact.HomeAddress != null)
            {
                address = new StructuredPostalAddress();
                address.FormattedAddress = oContact.HomeAddress;
                address.City = oContact.HomeAddressCity;
                address.Country=oContact.HomeAddressCountry ;
                address.Postcode=oContact.HomeAddressPostalCode ;
                address.Pobox=oContact.HomeAddressPostOfficeBox ;
                address.Subregion=oContact.HomeAddressState;
                address.Street=oContact.HomeAddressStreet;
                address.Rel = ContactsRelationships.IsHome;
                gContact.PostalAddresses.Add(address);
            }
            if (oContact.BusinessAddress != null)
            {
                address = new StructuredPostalAddress();
                address.FormattedAddress = oContact.BusinessAddress;
                address.City = oContact.BusinessAddressCity;
                address.Country = oContact.BusinessAddressCountry;
                address.Postcode = oContact.BusinessAddressPostalCode;
                address.Pobox = oContact.BusinessAddressPostOfficeBox;
                address.Subregion = oContact.BusinessAddressState;
                address.Street = oContact.BusinessAddressStreet;
                address.Rel = ContactsRelationships.IsWork;
                gContact.PostalAddresses.Add(address);
            }
            if (oContact.OtherAddress != null)
            {
                address = new StructuredPostalAddress();
                address.FormattedAddress = oContact.OtherAddress;
                address.City = oContact.OtherAddressCity;
                address.Country = oContact.OtherAddressCountry;
                address.Postcode = oContact.OtherAddressPostalCode;
                address.Pobox = oContact.OtherAddressPostOfficeBox;
                address.Subregion = oContact.OtherAddressState;
                address.Street = oContact.OtherAddressStreet;
                address.Rel = ContactsRelationships.IsOther;
                gContact.PostalAddresses.Add(address);
            }
            
            if(oContact.FullName!=null) {
                gContact.Title = oContact.FullName.ToString();
                gContact.Name.FullName = oContact.FullName.ToString(); 
            }
            else{
                gContact.Title="";
                gContact.Name=null;
            }

            gContact.ContactEntry.Birthday = oContact.Birthday != _outlookNullDateValue ? String.Format("{0:yyyy-MM-dd}", oContact.Birthday): null;

            gContact.Content = oContact.Body != null ? oContact.Body.ToString() : gContact.Content = null;
            if (oContact.CompanyName != null)
            {
                organ = new Organization();
                organ.Name = oContact.CompanyName;
                if (oContact.Title != null) { organ.Title = oContact.JobTitle; } else { organ.Title = null; }
                organ.Rel = ContactsRelationships.IsWork;
                gContact.Organizations.Add(organ);
            }
            
            if (oContact.IMAddress!=null) {
                imAddress = new IMAddress(oContact.IMAddress.ToString());
                imAddress.Rel = ContactsRelationships.IsOther;
                gContact.IMs.Add(imAddress); 
            }
            if (oContact.Initials != null) { gContact.ContactEntry.Initials = oContact.Initials.ToString(); } else { gContact.ContactEntry.Initials = null; }
            if (oContact.NickName!=null) { gContact.ContactEntry.Nickname = oContact.NickName.ToString(); }
            
            if (oContact.WebPage != null) {
                web = new Website();
                web.Href = oContact.WebPage;
                web.Label = oContact.WebPage;
                //web.Rel = ContactsRelationships.IsOther;
                gContact.ContactEntry.Websites.Add(web); 
            }

            //Emails
            if (oContact.Email1Address != null && oContact.Email1Address!="")
            {
                email = new EMail(oContact.Email1Address.ToString());
                email.Rel = ContactsRelationships.IsOther;
                if (oContact.Email1Address == oContact.Account)
                {
                    email.Primary = true;
                }
                gContact.Emails.Add(email);
            }
            if (oContact.Email2Address != null && oContact.Email2Address != "")
            {
                email = new EMail(oContact.Email2Address.ToString());
                email.Rel = ContactsRelationships.IsOther;
                if (oContact.Email2Address == oContact.Account)
                {
                    email.Primary = true;
                }
                gContact.Emails.Add(email);
            }
            if (oContact.Email3Address != null && oContact.Email3Address != "")
            {
                email = new EMail(oContact.Email3Address.ToString());
                email.Rel = ContactsRelationships.IsOther;
                if (oContact.Email3Address == oContact.Account)
                {
                    email.Primary = true;
                }
                gContact.Emails.Add(email);
            }
        }

        /// <summary>
        /// Find any attribute in extension collection property
        /// </summary>
        /// <param name="rel">Parameter to find</param>
        /// <param name="collection">Collection to search in</param>
        /// <returns>Return true if the parameter exists</returns>
        private bool FindGoogleRel(String rel, List<ICommonAttributes> collection)
        {
            foreach (ICommonAttributes attribute in collection)
            {
                if (attribute.Rel.Equals(rel))
                {
                    return true;
                }
            }
            return true;
        }

        /// <summary>
        /// Find a group name into main groups list
        /// </summary>
        /// <param name="groupName">Group name to search in</param>
        /// <returns>Group founded</returns>
        private Group FindGroup(string groupName)
        {
            Group group;
            foreach (DictionaryEntry groupEntry in ContactsGroupsList)
            {
                group = (Group)groupEntry.Value;
                if (groupName.Equals(group.Title))
                {
                    return group;
                }
            }
            return null;
        }

        /// <summary>
        /// Find a group into the group list of a google contact
        /// </summary>
        /// <param name="contact">Google contact object</param>
        /// <param name="groupName">Group name to search in</param>
        /// <returns>Group founded</returns>
        private Group FindGroupInContact(Contact contact,string groupName)
        {
            foreach (Google.GData.Contacts.GroupMembership msGroup in contact.GroupMembership)
            {
                Group group = (Group)ContactsGroupsList[msGroup.HRef];
                if (group.Title.Equals(groupName))
                {
                    return group;
                }
            }
            return null;
        }        
    }
}
