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
    /// Represents a file attachment.
    /// </summary>
    public sealed class FileAttachment : Attachment
    {
        private string fileName;
        private Stream contentStream;
        private byte[] content;
        private Stream loadToStream;
        private bool isContactPhoto;

        /// <summary>
        /// Initializes a new instance of the <see cref="FileAttachment"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal FileAttachment(Item owner)
            : base(owner)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FileAttachment"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal FileAttachment(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.FileAttachment;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="attachmentIndex">Index of this attachment.</param>
        internal override void Validate(int attachmentIndex)
        {
            if (string.IsNullOrEmpty(this.fileName) && (this.content == null) && (this.contentStream == null))
            {
                throw new ServiceValidationException(string.Format(Strings.FileAttachmentContentIsNotSet, attachmentIndex));
            }
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
                if (reader.LocalName == XmlElementNames.IsContactPhoto)
                {
                    this.isContactPhoto = reader.ReadElementValue<bool>();
                }
                else if (reader.LocalName == XmlElementNames.Content)
                {
                    if (this.loadToStream != null)
                    {
                        reader.ReadBase64ElementValue(this.loadToStream);
                    }
                    else
                    {
                        // If there's a file attachment content handler, use it. Otherwise
                        // load the content into a byte array.
                        // TODO: Should we mark the attachment to indicate that content is stored elsewhere?
                        if (reader.Service.FileAttachmentContentHandler != null)
                        {
                            Stream outputStream = reader.Service.FileAttachmentContentHandler.GetOutputStream(this.Id);

                            if (outputStream != null)
                            {
                                reader.ReadBase64ElementValue(outputStream);
                            }
                            else
                            {
                                this.content = reader.ReadBase64ElementValue();
                            }
                        }
                        else
                        {
                            this.content = reader.ReadBase64ElementValue();
                        }
                    }

                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// For FileAttachment, the only thing need to patch is the AttachmentId.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal override bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
        {
            return base.TryReadElementFromXml(reader);
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            base.LoadFromJson(jsonProperty, service);

            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.IsContactPhoto:
                        this.isContactPhoto = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.Content:
                        if (this.loadToStream != null)
                        {
                            jsonProperty.ReadAsBase64Content(key, this.loadToStream);
                        }
                        else
                        {
                            // If there's a file attachment content handler, use it. Otherwise
                            // load the content into a byte array.
                            // TODO: Should we mark the attachment to indicate that content is stored elsewhere?
                            if (service.FileAttachmentContentHandler != null)
                            {
                                Stream outputStream = service.FileAttachmentContentHandler.GetOutputStream(this.Id);

                                if (outputStream != null)
                                {
                                    jsonProperty.ReadAsBase64Content(key, outputStream);
                                }
                                else
                                {
                                    this.content = jsonProperty.ReadAsBase64Content(key);
                                }
                            }
                            else
                            {
                                this.content = jsonProperty.ReadAsBase64Content(key);
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements and content to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsContactPhoto, this.isContactPhoto);
            }

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Content);

            if (!string.IsNullOrEmpty(this.FileName))
            {
                using (FileStream fileStream = new FileStream(this.FileName, FileMode.Open, FileAccess.Read))
                {
                    writer.WriteBase64ElementValue(fileStream);
                }
            }
            else if (this.ContentStream != null)
            {
                writer.WriteBase64ElementValue(this.ContentStream);
            }
            else if (this.Content != null)
            {
                writer.WriteBase64ElementValue(this.Content);
            }
            else
            {
                EwsUtilities.Assert(
                    false,
                    "FileAttachment.WriteElementsToXml",
                    "The attachment's content is not set.");
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonAttachment = base.InternalToJson(service) as JsonObject;

            if (service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                jsonAttachment.Add(XmlElementNames.IsContactPhoto, this.isContactPhoto);
            }
            
            if (!string.IsNullOrEmpty(this.FileName))
            {
                using (FileStream fileStream = new FileStream(this.FileName, FileMode.Open, FileAccess.Read))
                {
                    jsonAttachment.AddBase64(XmlElementNames.Content, fileStream);
                }
            }
            else if (this.ContentStream != null)
            {
                jsonAttachment.AddBase64(XmlElementNames.Content, this.ContentStream);
            }
            else if (this.Content != null)
            {
                jsonAttachment.AddBase64(XmlElementNames.Content, this.Content);
            }
            else
            {
                EwsUtilities.Assert(
                    false,
                    "FileAttachment.WriteElementsToXml",
                    "The attachment's content is not set.");
            }

            return jsonAttachment;
        }

        /// <summary>
        /// Loads the content of the file attachment into the specified stream. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="stream">The stream to load the content of the attachment into.</param>
        public void Load(Stream stream)
        {
            this.loadToStream = stream;

            try
            {
                this.Load();
            }
            finally
            {
                this.loadToStream = null;
            }
        }

        /// <summary>
        /// Loads the content of the file attachment into the specified file. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="fileName">The name of the file to load the content of the attachment into. If the file already exists, it is overwritten.</param>
        public void Load(string fileName)
        {
            this.loadToStream = new FileStream(fileName, FileMode.Create);

            try
            {
                this.Load();
            }
            finally
            {
                this.loadToStream.Dispose();
                this.loadToStream = null;
            }

            this.fileName = fileName;
            this.content = null;
            this.contentStream = null;
        }

        /// <summary>
        /// Gets the name of the file the attachment is linked to.
        /// </summary>
        public string FileName
        {
            get
            {
                return this.fileName;
            }

            internal set
            {
                this.ThrowIfThisIsNotNew();

                this.fileName = value;
                this.content = null;
                this.contentStream = null;
            }
        }

        /// <summary>
        /// Gets or sets the content stream.
        /// </summary>
        /// <value>The content stream.</value>
        internal Stream ContentStream
        {
            get
            {
                return this.contentStream;
            }

            set
            {
                this.ThrowIfThisIsNotNew();
                
                this.contentStream = value;
                this.content = null;
                this.fileName = null;
            }
        }

        /// <summary>
        /// Gets the content of the attachment into memory. Content is set only when Load() is called.
        /// </summary>
        public byte[] Content
        {
            get
            {
                return this.content;
            }

            internal set
            {
                this.ThrowIfThisIsNotNew();
                
                this.content = value;
                this.fileName = null;
                this.contentStream = null;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this attachment is a contact photo.
        /// </summary>
        public bool IsContactPhoto
        {
            get 
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "IsContactPhoto");

                return this.isContactPhoto;
            }

            set
            {
                EwsUtilities.ValidatePropertyVersion(this.Service, ExchangeVersion.Exchange2010, "IsContactPhoto");

                this.ThrowIfThisIsNotNew();

                this.isContactPhoto = value;
            }
        }
    }
}