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

    /// <summary>
    /// Represents a GetAttachment request.
    /// </summary>
    internal sealed class GetAttachmentRequest : MultiResponseServiceRequest<GetAttachmentResponse>
    {
        private List<Attachment> attachments = new List<Attachment>();
        private List<string> attachmentIds = new List<string>();
        private List<PropertyDefinitionBase> additionalProperties = new List<PropertyDefinitionBase>();
        private BodyType? bodyType;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetAttachmentRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal GetAttachmentRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            if (this.Attachments.Count > 0)
            {
                EwsUtilities.ValidateParamCollection(this.Attachments, "Attachments");
            }

            if (this.AttachmentIds.Count > 0)
            {
                EwsUtilities.ValidateParamCollection(this.AttachmentIds, "AttachmentIds");
            }

            if (this.AttachmentIds.Count == 0 && this.Attachments.Count == 0)
            {
                throw new ArgumentException(Strings.CollectionIsEmpty, @"Attachments/AttachmentIds");
            }
            for (int i = 0; i < this.AdditionalProperties.Count; i++)
            {
                EwsUtilities.ValidateParam(this.AdditionalProperties[i], string.Format("AdditionalProperties[{0}]", i));
            }
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override GetAttachmentResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new GetAttachmentResponse(this.Attachments.Count > 0 ? this.Attachments[responseIndex] : null);
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.Attachments.Count + this.AttachmentIds.Count;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetAttachment;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetAttachmentResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.GetAttachmentResponseMessage;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.BodyType.HasValue || this.AdditionalProperties.Count > 0)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.AttachmentShape);

                if (this.BodyType.HasValue)
                {
                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.BodyType,
                        this.BodyType.Value);
                }

                if (this.AdditionalProperties.Count > 0)
                {
                    PropertySet.WriteAdditionalPropertiesToXml(writer, this.AdditionalProperties);
                }

                writer.WriteEndElement(); // AttachmentShape
            }

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.AttachmentIds);

            foreach (Attachment attachment in this.Attachments)
            {
                this.WriteAttachmentIdXml(writer, attachment.Id);
            }

            foreach (string attachmentId in this.AttachmentIds)
            {
                this.WriteAttachmentIdXml(writer, attachmentId);
            }

            writer.WriteEndElement();
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

        /// <summary>
        /// Gets the attachment ids.
        /// </summary>
        /// <value>The attachment ids.</value>
        public List<string> AttachmentIds
        {
            get { return this.attachmentIds; }
        }

        /// <summary>
        /// Gets the additional properties.
        /// </summary>
        /// <value>The additional properties.</value>
        public List<PropertyDefinitionBase> AdditionalProperties
        {
            get { return this.additionalProperties; }
        }

        /// <summary>
        /// Gets or sets the type of the body.
        /// </summary>
        /// <value>The type of the body.</value>
        public BodyType? BodyType
        {
            get { return this.bodyType; }
            set { this.bodyType = value; }
        }

        /// <summary>
        /// Gets a value indicating whether the TimeZoneContext SOAP header should be emitted.
        /// </summary>
        /// <value>
        ///     <c>true</c> if the time zone should be emitted; otherwise, <c>false</c>.
        /// </value>
        internal override bool EmitTimeZoneHeader
        {
            get
            {
                // we currently do not emit "AttachmentResponseShapeType.IncludeMimeContent"
                //
                return this.additionalProperties.Contains(ItemSchema.MimeContent);
            }
        }

        /// <summary>
        /// Writes attachment id elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="attachmentId">The attachment id.</param>
        private void WriteAttachmentIdXml(EwsServiceXmlWriter writer, string attachmentId)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.AttachmentId);
            writer.WriteAttributeValue(XmlAttributeNames.Id, attachmentId);
            writer.WriteEndElement();
        }
    }
}