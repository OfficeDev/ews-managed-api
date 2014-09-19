// ---------------------------------------------------------------------------
// <copyright file="ResponseMessage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseMessage class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the base class for e-mail related responses (Reply, Reply all and Forward).
    /// </summary>
    public sealed class ResponseMessage : ResponseObject<EmailMessage>
    {
        private ResponseMessageType responseType;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResponseMessage"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        /// <param name="responseType">Type of the response.</param>
        internal ResponseMessage(Item referenceItem, ResponseMessageType responseType)
            : base(referenceItem)
        {
            this.responseType = responseType;
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return ResponseMessageSchema.Instance;
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// This methods lets subclasses of ServiceObject override the default mechanism
        /// by which the XML element name associated with their type is retrieved.
        /// </summary>
        /// <returns>
        /// The XML element name associated with this type.
        /// If this method returns null or empty, the XML element name associated with this
        /// type is determined by the EwsObjectDefinition attribute that decorates the type,
        /// if present.
        /// </returns>
        /// <remarks>
        /// Item and folder classes that can be returned by EWS MUST rely on the EwsObjectDefinition
        /// attribute for XML element name determination.
        /// </remarks>
        internal override string GetXmlElementNameOverride()
        {
            switch (this.responseType)
            {
                case ResponseMessageType.Reply:
                    return XmlElementNames.ReplyToItem;
                case ResponseMessageType.ReplyAll:
                    return XmlElementNames.ReplyAllToItem;
                case ResponseMessageType.Forward:
                    return XmlElementNames.ForwardItem;
                default:
                    EwsUtilities.Assert(
                        false,
                        "ResponseMessage.GetXmlElementNameOverride",
                        "An unexpected value for responseType could not be handled.");
                    return null; // Because the compiler wants it
            }
        }

        /// <summary>
        /// Gets a value indicating the type of response this object represents.
        /// </summary>
        public ResponseMessageType ResponseType
        {
            get { return this.responseType; }
        }

        #region Properties

        /// <summary>
        /// Gets or sets the body of the response.
        /// </summary>
        public MessageBody Body
        {
            get { return (MessageBody)this.PropertyBag[ItemSchema.Body]; }
            set { this.PropertyBag[ItemSchema.Body] = value; }
        }

        /// <summary>
        /// Gets a list of recipients the response will be sent to.
        /// </summary>
        public EmailAddressCollection ToRecipients
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.ToRecipients]; }
        }

        /// <summary>
        /// Gets a list of recipients the response will be sent to as Cc.
        /// </summary>
        public EmailAddressCollection CcRecipients
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.CcRecipients]; }
        }

        /// <summary>
        /// Gets a list of recipients this response will be sent to as Bcc.
        /// </summary>
        public EmailAddressCollection BccRecipients
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.BccRecipients]; }
        }

        /// <summary>
        /// Gets or sets the subject of this response.
        /// </summary>
        public string Subject
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.Subject]; }
            set { this.PropertyBag[EmailMessageSchema.Subject] = value; }
        }

        /// <summary>
        /// Gets or sets the body prefix of this response. The body prefix will be prepended to the original
        /// message's body when the response is created.
        /// </summary>
        public MessageBody BodyPrefix
        {
            get { return (MessageBody)this.PropertyBag[ResponseObjectSchema.BodyPrefix]; }
            set { this.PropertyBag[ResponseObjectSchema.BodyPrefix] = value; }
        }
        #endregion
    }
}
