// ---------------------------------------------------------------------------
// <copyright file="SearchableMailbox.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchableMailbox class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents searchable mailbox object
    /// </summary>
    public sealed class SearchableMailbox
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public SearchableMailbox()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="guid">Guid</param>
        /// <param name="smtpAddress">Smtp address</param>
        /// <param name="isExternalMailbox">If true, this is an external mailbox</param>
        /// <param name="externalEmailAddress">External email address</param>
        /// <param name="displayName">Display name</param>
        /// <param name="isMembershipGroup">Is a membership group</param>
        /// <param name="referenceId">Reference id</param>
        public SearchableMailbox(
            Guid guid, 
            string smtpAddress, 
            bool isExternalMailbox,
            string externalEmailAddress,
            string displayName, 
            bool isMembershipGroup, 
            string referenceId)
        {
            this.Guid = guid;
            this.SmtpAddress = smtpAddress;
            this.IsExternalMailbox = isExternalMailbox;
            this.ExternalEmailAddress = externalEmailAddress;
            this.DisplayName = displayName;
            this.IsMembershipGroup = isMembershipGroup;
            this.ReferenceId = referenceId;
        }

        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Searchable mailbox object</returns>
        internal static SearchableMailbox LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.SearchableMailbox);

            SearchableMailbox searchableMailbox = new SearchableMailbox();
            searchableMailbox.Guid = new Guid(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Guid));
            searchableMailbox.SmtpAddress = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.PrimarySmtpAddress);
            bool isExternalMailbox = false;
            bool.TryParse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.IsExternalMailbox), out isExternalMailbox);
            searchableMailbox.IsExternalMailbox = isExternalMailbox;
            searchableMailbox.ExternalEmailAddress = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ExternalEmailAddress);
            searchableMailbox.DisplayName = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.DisplayName);
            bool isMembershipGroup = false;
            bool.TryParse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.IsMembershipGroup), out isMembershipGroup);
            searchableMailbox.IsMembershipGroup = isMembershipGroup;
            searchableMailbox.ReferenceId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ReferenceId);

            return searchableMailbox;
        }

        /// <summary>
        /// Load from json
        /// </summary>
        /// <param name="jsonObject">The json object</param>
        /// <returns>Searchable mailbox object</returns>
        internal static SearchableMailbox LoadFromJson(JsonObject jsonObject)
        {
            SearchableMailbox searchableMailbox = new SearchableMailbox();

            if (jsonObject.ContainsKey(XmlElementNames.Guid))
            {
                searchableMailbox.Guid = new Guid(jsonObject.ReadAsString(XmlElementNames.Guid));
            }

            if (jsonObject.ContainsKey(XmlElementNames.DisplayName))
            {
                searchableMailbox.DisplayName = jsonObject.ReadAsString(XmlElementNames.DisplayName);
            }

            if (jsonObject.ContainsKey(XmlElementNames.PrimarySmtpAddress))
            {
                searchableMailbox.SmtpAddress = jsonObject.ReadAsString(XmlElementNames.PrimarySmtpAddress);
            }

            if (jsonObject.ContainsKey(XmlElementNames.IsExternalMailbox))
            {
                searchableMailbox.IsExternalMailbox = jsonObject.ReadAsBool(XmlElementNames.IsExternalMailbox);
            }

            if (jsonObject.ContainsKey(XmlElementNames.ExternalEmailAddress))
            {
                searchableMailbox.ExternalEmailAddress = jsonObject.ReadAsString(XmlElementNames.ExternalEmailAddress);
            }

            if (jsonObject.ContainsKey(XmlElementNames.IsMembershipGroup))
            {
                searchableMailbox.IsMembershipGroup = jsonObject.ReadAsBool(XmlElementNames.IsMembershipGroup);
            }

            if (jsonObject.ContainsKey(XmlElementNames.ReferenceId))
            {
                searchableMailbox.ReferenceId = jsonObject.ReadAsString(XmlElementNames.ReferenceId);
            }

            return searchableMailbox;
        }

        /// <summary>
        /// Guid
        /// </summary>
        public Guid Guid { get; set; }

        /// <summary>
        /// Smtp address
        /// </summary>
        public string SmtpAddress { get; set; }

        /// <summary>
        /// If true, this is an external mailbox
        /// </summary>
        public bool IsExternalMailbox { get; set; }

        /// <summary>
        /// External email address for the mailbox
        /// </summary>
        public string ExternalEmailAddress { get; set; }

        /// <summary>
        /// Display name
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Is a membership group
        /// </summary>
        public bool IsMembershipGroup { get; set; }

        /// <summary>
        /// Reference id
        /// </summary>
        public string ReferenceId { get; set; }
    }
}
