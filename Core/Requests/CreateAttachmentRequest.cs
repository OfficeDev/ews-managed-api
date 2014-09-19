// ---------------------------------------------------------------------------
// <copyright file="CreateAttachmentRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateAttachmentRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a CreateAttachment request.
    /// </summary>
    internal sealed class CreateAttachmentRequest : MultiResponseServiceRequest<CreateAttachmentResponse>, IJsonSerializable
    {
        private string parentItemId;
        private List<Attachment> attachments = new List<Attachment>();

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateAttachmentRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal CreateAttachmentRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request..
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.ParentItemId, "ParentItemId");
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override CreateAttachmentResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new CreateAttachmentResponse(this.Attachments[responseIndex]);
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.Attachments.Count;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.CreateAttachment;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.CreateAttachmentResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.CreateAttachmentResponseMessage;
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ParentItemId);
            writer.WriteAttributeValue(XmlAttributeNames.Id, this.ParentItemId);
            writer.WriteEndElement();

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Attachments);
            foreach (Attachment attachment in this.Attachments)
            {
                attachment.WriteToXml(writer, attachment.GetXmlElementName());
            }
            writer.WriteEndElement();
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonRequest = new JsonObject();

            jsonRequest.Add(XmlElementNames.ParentItemId, new ItemId(this.ParentItemId).InternalToJson(service));

            List<object> attachmentArray = new List<object>();
            foreach (Attachment attachment in this.Attachments)
            {
                attachmentArray.Add(attachment.InternalToJson(service));
            }

            jsonRequest.Add(XmlElementNames.Attachments, attachmentArray.ToArray());

            return jsonRequest;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Gets a value indicating whether the TimeZoneContext SOAP header should be emitted.
        /// </summary>
        internal override bool EmitTimeZoneHeader
        {
            get
            {
                foreach (ItemAttachment itemAttachment in this.attachments.OfType<ItemAttachment>())
                {
                    if ((itemAttachment.Item != null) && itemAttachment.Item.GetIsTimeZoneHeaderRequired(false /* isUpdateOperation */))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

        /// <summary>
        /// Gets the attachments.
        /// </summary>
        /// <value>The attachments.</value>
        public List<Attachment> Attachments
        {
            get { return this.attachments; }
        }

        /// <summary>
        /// Gets or sets the parent item id.
        /// </summary>
        /// <value>The parent item id.</value>
        public string ParentItemId
        {
            get { return this.parentItemId; }
            set { this.parentItemId = value; }
        }
    }
}
