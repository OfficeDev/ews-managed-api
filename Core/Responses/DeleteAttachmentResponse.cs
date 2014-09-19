// ---------------------------------------------------------------------------
// <copyright file="DeleteAttachmentResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeleteAttachmentResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual attachment deletion operation.
    /// </summary>
    public sealed class DeleteAttachmentResponse : ServiceResponse
    {
        private Attachment attachment;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteAttachmentResponse"/> class.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        internal DeleteAttachmentResponse(Attachment attachment)
            : base()
        {
            EwsUtilities.Assert(
                attachment != null,
                "DeleteAttachmentResponse.ctor",
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

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RootItemId);

            string changeKey = reader.ReadAttributeValue(XmlAttributeNames.RootItemChangeKey);
            if (!string.IsNullOrEmpty(changeKey) && this.attachment.Owner != null)
            {
                this.attachment.Owner.RootItemId.ChangeKey = changeKey;
            }

            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.RootItemId);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            if (responseObject.ContainsKey(XmlElementNames.RootItemId))
            {
                JsonObject jsonRootItemId = responseObject.ReadAsJsonObject(XmlElementNames.RootItemId);
                string changeKey;

                if (jsonRootItemId.ContainsKey(XmlAttributeNames.RootItemChangeKey) &&
                    !String.IsNullOrEmpty(changeKey = jsonRootItemId.ReadAsString(XmlAttributeNames.RootItemChangeKey)) &&
                    this.attachment.Owner != null)
                {
                    this.attachment.Owner.RootItemId.ChangeKey = changeKey;
                }
            }
        }

        /// <summary>
        /// Gets the attachment that was deleted.
        /// </summary>
        internal Attachment Attachment
        {
            get { return this.attachment; }
        }
    }
}
