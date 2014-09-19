// ---------------------------------------------------------------------------
// <copyright file="CreateAttachmentResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateAttachmentResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents the response to an individual attachment creation operation.
    /// </summary>
    public sealed class CreateAttachmentResponse : ServiceResponse
    {
        private Attachment attachment;

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateAttachmentResponse"/> class.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        internal CreateAttachmentResponse(Attachment attachment)
            : base()
        {
            EwsUtilities.Assert(
                attachment != null,
                "CreateAttachmentResponse.ctor",
                "attachment is null");

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

            reader.Read(XmlNodeType.Element);
            this.attachment.LoadFromXml(reader, reader.LocalName);

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.Attachments);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            object[] attachmentArray = responseObject.ReadAsArray(XmlElementNames.Attachments);

            if (attachmentArray != null && attachmentArray.Length > 0)
            {
                this.attachment.LoadFromJson(attachmentArray[0] as JsonObject, service);
            }
        }

        /// <summary>
        /// Gets the attachment that was created.
        /// </summary>
        internal Attachment Attachment
        {
            get { return this.attachment; }
        }
    }
}
