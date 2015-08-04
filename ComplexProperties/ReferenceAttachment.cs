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

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Represents an attachment by reference.
    /// </summary>
    public sealed class ReferenceAttachment : Attachment
    {
        /// <summary>
        /// The AttachLongPathName of the attachment.
        /// </summary>
        private string attachLongPathName;

        /// <summary>
        /// The ProviderType of the attachment.
        /// </summary>
        private string providerType;

        /// <summary>
        /// The ProviderEndpointUrl of the attachment.
        /// </summary>
        private string providerEndpointUrl;

        /// <summary>
        /// The AttachmentThumbnailUrl of the attachment.
        /// </summary>
        private string attachmentThumbnailUrl;

        /// <summary>
        /// The AttachmentPreviewUrl of the attachment.
        /// </summary>
        private string attachmentPreviewUrl;

        /// <summary>
        /// The PermissionType of the attachment.
        /// </summary>
        private int permissionType;

        /// <summary>
        /// The AttachmentIsFolder of the attachment.
        /// </summary>
        private bool attachmentIsFolder;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReferenceAttachment"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal ReferenceAttachment(Item owner)
            : base(owner)
        {
            EwsUtilities.ValidateClassVersion(
                this.Owner.Service,
                ExchangeVersion.Exchange2015,
                this.GetType().Name);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ReferenceAttachment;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            bool result = base.TryReadElementFromXml(reader);

            if (!result)
            {
                if (reader.LocalName == XmlElementNames.AttachLongPathName)
                {
                    this.attachLongPathName = reader.ReadElementValue();
                    return true;
                }

                if (reader.LocalName == XmlElementNames.ProviderType)
                {
                    this.providerType = reader.ReadElementValue();
                    return true;
                }

                if (reader.LocalName == XmlElementNames.ProviderEndpointUrl)
                {
                    this.providerEndpointUrl = reader.ReadElementValue();
                    return true;
                }

                if (reader.LocalName == XmlElementNames.AttachmentThumbnailUrl)
                {
                    this.attachmentThumbnailUrl = reader.ReadElementValue();
                    return true;
                }

                if (reader.LocalName == XmlElementNames.AttachmentPreviewUrl)
                {
                    this.attachmentPreviewUrl = reader.ReadElementValue();
                    return true;
                }

                if (reader.LocalName == XmlElementNames.PermissionType)
                {
                    this.permissionType = reader.ReadElementValue<int>();
                    return true;
                }

                if (reader.LocalName == XmlElementNames.AttachmentIsFolder)
                {
                    this.attachmentIsFolder = reader.ReadElementValue<bool>();
                    return true;
                }
            }

            return result;
        }

        /// <summary>
        /// For ReferenceAttachment, the only thing need to patch is the AttachmentId.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
        {
            return base.TryReadElementFromXml(reader);
        }

        /// <summary>
        /// Writes elements and content to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AttachLongPathName, this.AttachLongPathName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ProviderType, this.ProviderType);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ProviderEndpointUrl, this.ProviderEndpointUrl);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AttachmentThumbnailUrl, this.AttachmentThumbnailUrl);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AttachmentPreviewUrl, this.AttachmentPreviewUrl);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PermissionType, this.PermissionType);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AttachmentIsFolder, this.AttachmentIsFolder);
        }

        /// <summary>
        /// Gets or sets a fully-qualified path identifying the attachment.
        /// </summary>
        public string AttachLongPathName
        {
            get { return this.attachLongPathName; }
            set { this.SetFieldValue<string>(ref this.attachLongPathName, value); }
        }

        /// <summary>
        /// Gets or sets the type of the attachment provider.
        /// </summary>
        public string ProviderType
        {
            get { return this.providerType; }
            set { this.SetFieldValue<string>(ref this.providerType, value); }
        }

        /// <summary>
        /// Gets or sets the URL of the attachment provider.
        /// </summary>
        public string ProviderEndpointUrl
        {
            get { return this.providerEndpointUrl; }
            set { this.SetFieldValue<string>(ref this.providerEndpointUrl, value); }
        }

        /// <summary>
        /// Gets or sets the URL of the attachment thumbnail.
        /// </summary>
        public string AttachmentThumbnailUrl
        {
            get { return this.attachmentThumbnailUrl; }
            set { this.SetFieldValue<string>(ref this.attachmentThumbnailUrl, value); }
        }

        /// <summary>
        /// Gets or sets the URL of the attachment preview.
        /// </summary>
        public string AttachmentPreviewUrl
        {
            get { return this.attachmentPreviewUrl; }
            set { this.SetFieldValue<string>(ref this.attachmentPreviewUrl, value); }
        }

        /// <summary>
        /// Gets or sets the permission of the attachment.
        /// </summary>
        public int PermissionType
        {
            get { return this.permissionType; }
            set { this.SetFieldValue<int>(ref this.permissionType, value); }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the attachment points to a folder.
        /// </summary>
        public bool AttachmentIsFolder
        {
            get { return this.attachmentIsFolder; }
            set { this.SetFieldValue<bool>(ref this.attachmentIsFolder, value); }
        }
    }
}
