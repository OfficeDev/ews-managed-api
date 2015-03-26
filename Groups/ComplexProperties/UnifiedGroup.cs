// ---------------------------------------------------------------------------
// <copyright file="UnifiedGroup.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UnifiedGroup class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a UnifiedGroup class.
    /// </summary>
    public class UnifiedGroup : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UnifiedGroup"/> class.
        /// </summary>
        internal UnifiedGroup() :
            base()
        {
        }

        /// <summary>
        /// Gets or sets whether this groups is a favorite group
        /// </summary>
        public bool IsFavorite { get; set; }

        /// <summary>
        /// Gets or sets the ExternalDirectoryObjectId for this group
        /// </summary>
        public string ExternalDirectoryObjectId { get; set; }

        /// <summary>
        /// Gets or sets the LastVisitedTimeUtc for this group and user
        /// </summary>
        public string LastVisitedTimeUtc { get; set; }

        /// <summary>
        /// Gets or sets the SmtpAddress associated with this group
        /// </summary>
        public string SmtpAddress { get; set; }

        /// <summary>
        /// Gets or sets the LegacyDN associated with this group
        /// </summary>
        public string LegacyDN { get; set; }

        /// <summary>
        /// Gets or sets the MailboxGuid associated with this group
        /// </summary>
        public string MailboxGuid { get; set; }

        /// <summary>
        /// Gets or sets the DisplayName associated with this group
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets the AccessType associated with this group
        /// </summary>
        public UnifiedGroupAccessType AccessType { get; set; }

        /// <summary>
        /// Read Conversations from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">The xml element to read.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.UnifiedGroup);
            do
            {
                reader.Read();
                switch (reader.LocalName)
                {
                    case XmlElementNames.SmtpAddress:
                        this.SmtpAddress = reader.ReadElementValue();
                        break;
                    case XmlElementNames.LegacyDN:
                        this.LegacyDN = reader.ReadElementValue();
                        break;
                    case XmlElementNames.MailboxGuid:
                        this.MailboxGuid = reader.ReadElementValue();
                        break;
                    case XmlElementNames.DisplayName:
                        this.DisplayName = reader.ReadElementValue();
                        break;
                    case XmlElementNames.IsFavorite:
                        this.IsFavorite = reader.ReadElementValue<bool>();
                        break;
                    case XmlElementNames.LastVisitedTimeUtc:
                        this.LastVisitedTimeUtc = reader.ReadElementValue();
                        break;
                    case XmlElementNames.AccessType:
                        this.AccessType = (UnifiedGroupAccessType)Enum.Parse(typeof(UnifiedGroupAccessType), reader.ReadElementValue(), false);
                        break;
                    case XmlElementNames.ExternalDirectoryObjectId:
                        this.ExternalDirectoryObjectId = reader.ReadElementValue();
                        break;
                    default:
                        break;
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.UnifiedGroup));

            // Skip end element
            reader.EnsureCurrentNodeIsEndElement(XmlNamespace.NotSpecified, XmlElementNames.UnifiedGroup);
            reader.Read(); 
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject responseObject, ExchangeService service)
        {
            if (responseObject.ContainsKey(XmlElementNames.SmtpAddress))
            {
                this.SmtpAddress = responseObject.ReadAsString(XmlElementNames.SmtpAddress);
            }

            if (responseObject.ContainsKey(XmlElementNames.LegacyDN))
            {
                this.LegacyDN = responseObject.ReadAsString(XmlElementNames.LegacyDN);
            }

            if (responseObject.ContainsKey(XmlElementNames.MailboxGuid))
            {
                this.MailboxGuid = responseObject.ReadAsString(XmlElementNames.MailboxGuid);
            }

            if (responseObject.ContainsKey(XmlElementNames.DisplayName))
            {
                this.DisplayName = responseObject.ReadAsString(XmlElementNames.DisplayName);
            }

            if (responseObject.ContainsKey(XmlElementNames.IsFavorite))
            {
                this.IsFavorite = responseObject.ReadAsBool(XmlElementNames.IsFavorite);
            }

            if (responseObject.ContainsKey(XmlElementNames.LastVisitedTimeUtc))
            {
                this.LastVisitedTimeUtc = responseObject.ReadAsString(XmlElementNames.LastVisitedTimeUtc);
            }

            if (responseObject.ContainsKey(XmlElementNames.AccessType))
            {
                this.AccessType = (UnifiedGroupAccessType)Enum.Parse(typeof(UnifiedGroupAccessType), responseObject.ReadAsString(XmlElementNames.AccessType), false);
            }

            if (responseObject.ContainsKey(XmlElementNames.ExternalDirectoryObjectId))
            {
                this.ExternalDirectoryObjectId = responseObject.ReadAsString(XmlElementNames.ExternalDirectoryObjectId);
            }
        }
    }
}