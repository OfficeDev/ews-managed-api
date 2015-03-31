// ---------------------------------------------------------------------------
// <copyright file="GetAttachmentResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetAttachmentResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents the response to an individual attachment retrieval request.
    /// </summary>
    public sealed class GetAttachmentResponse : ServiceResponse
    {
        private Attachment attachment;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetAttachmentResponse"/> class.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        internal GetAttachmentResponse(Attachment attachment)
            : base()
        {
            this.attachment = attachment;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Attachments);
            if (!reader.IsEmptyElement)
            {
                reader.Read(XmlNodeType.Element);

                if (this.attachment == null)
                {
                    if (string.Equals(reader.LocalName, XmlElementNames.FileAttachment, StringComparison.OrdinalIgnoreCase))
                    {
                        this.attachment = new FileAttachment(reader.Service);
                    }
                    else if (string.Equals(reader.LocalName, XmlElementNames.ItemAttachment, StringComparison.OrdinalIgnoreCase))
                    {
                        this.attachment = new ItemAttachment(reader.Service);
                    }
                    else if (string.Equals(reader.LocalName, XmlElementNames.ReferenceAttachment, StringComparison.OrdinalIgnoreCase))
                    {
                        this.attachment = new ReferenceAttachment(reader.Service);
                    }

                }

                if (this.attachment != null)
                {
                    this.attachment.LoadFromXml(reader, reader.LocalName);
                }

                reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.Attachments);
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            object[] attachmentsArray;
            if (responseObject.ContainsKey(XmlElementNames.Attachments) &&
                (attachmentsArray = responseObject.ReadAsArray(XmlElementNames.Attachments)).Length > 0)
            {
                JsonObject attachmentArrayJsonObject = attachmentsArray[0] as JsonObject;
                
                if (this.attachment == null && attachmentArrayJsonObject != null)
                {
                    if (attachmentArrayJsonObject.ContainsKey(XmlElementNames.FileAttachment))
                    {
                        this.attachment = new FileAttachment(service);
                    }
                    else if (attachmentArrayJsonObject.ContainsKey(XmlElementNames.ItemAttachment))
                    {
                        this.attachment = new ItemAttachment(service);
                    }
                    else if (attachmentArrayJsonObject.ContainsKey(XmlElementNames.ReferenceAttachment))
                    {
                        this.attachment = new ReferenceAttachment(service);
                    }
                }

                if (this.attachment != null)
                {
                    this.attachment.LoadFromJson(attachmentArrayJsonObject, service);
                }
            }
        }

        /// <summary>
        /// Gets the attachment that was retrieved.
        /// </summary>
        public Attachment Attachment
        {
            get { return this.attachment; }
        }
    }
}
