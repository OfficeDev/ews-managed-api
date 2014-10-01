#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the CalendarResponseMessage class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the base class for accept, tentatively accept and decline response messages.
    /// </summary>
    /// <typeparam name="TMessage">The type of message that is created when this response message is saved.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class CalendarResponseMessage<TMessage> : CalendarResponseMessageBase<TMessage>
        where TMessage : EmailMessage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarResponseMessage&lt;TMessage&gt;"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal CalendarResponseMessage(Item referenceItem)
            : base(referenceItem)
        {
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return CalendarResponseObjectSchema.Instance;
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

        // TODO : Does this need to be exposed?
        internal string ItemClass
        {
            get { return (string)this.PropertyBag[ItemSchema.ItemClass]; }
            set { this.PropertyBag[ItemSchema.ItemClass] = value; }
        }

        /// <summary>
        /// Gets or sets the sensitivity of this response.
        /// </summary>
        public Sensitivity Sensitivity
        {
            get { return (Sensitivity)this.PropertyBag[ItemSchema.Sensitivity]; }
            set { this.PropertyBag[ItemSchema.Sensitivity] = value; }
        }

        /// <summary>
        /// Gets a list of attachments to this response.
        /// </summary>
        public AttachmentCollection Attachments
        {
            get { return (AttachmentCollection)this.PropertyBag[ItemSchema.Attachments]; }
        }

        // TODO : Does this need to be exposed?
        internal InternetMessageHeaderCollection InternetMessageHeaders
        {
            get { return (InternetMessageHeaderCollection)this.PropertyBag[ItemSchema.InternetMessageHeaders]; }
        }

        /// <summary>
        /// Gets or sets the sender of this response.
        /// </summary>
        public EmailAddress Sender
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.Sender]; }
            set { this.PropertyBag[EmailMessageSchema.Sender] = value; }
        }

        #endregion
    }
}
