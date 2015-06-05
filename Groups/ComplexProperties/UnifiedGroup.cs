/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

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
    }
}