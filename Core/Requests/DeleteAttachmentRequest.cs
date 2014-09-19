// ---------------------------------------------------------------------------
// <copyright file="DeleteAttachmentRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeleteAttachmentRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a DeleteAttachment request.
    /// </summary>
    internal sealed class DeleteAttachmentRequest : MultiResponseServiceRequest<DeleteAttachmentResponse>, IJsonSerializable
    {
        private List<Attachment> attachments = new List<Attachment>();

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteAttachmentRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal DeleteAttachmentRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.Attachments, "Attachments");
            for (int i = 0; i < this.Attachments.Count; i++)
            {
                EwsUtilities.ValidateParam(this.Attachments[i].Id, string.Format("Attachment[{0}].Id", i));
            }
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service object.</returns>
        internal override DeleteAttachmentResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new DeleteAttachmentResponse(this.Attachments[responseIndex]);
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
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.DeleteAttachment;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.DeleteAttachmentResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.DeleteAttachmentResponseMessage;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.AttachmentIds);

            foreach (Attachment attachment in this.Attachments)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.AttachmentId);
                writer.WriteAttributeValue(XmlAttributeNames.Id, attachment.Id);
                writer.WriteEndElement();
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
            List<object> attachmentIds = new List<object>();
            
            foreach (Attachment attachment in this.Attachments)
            {
                JsonObject jsonAttachmentId = new JsonObject();
                jsonAttachmentId.AddTypeParameter("AttachmentId");
                jsonAttachmentId.Add(XmlAttributeNames.Id, attachment.Id);

                attachmentIds.Add(jsonAttachmentId);
            }

            jsonRequest.Add(XmlElementNames.AttachmentIds, attachmentIds.ToArray());

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
        /// Gets the attachments.
        /// </summary>
        /// <value>The attachments.</value>
        public List<Attachment> Attachments
        {
            get { return this.attachments; }
        }
    }
}
