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
    using System.Text;

    /// <summary>
    /// Represents a post item. Properties available on post items are defined in the PostItemSchema class.
    /// </summary>
    [Attachable]
    [ServiceObjectDefinition(XmlElementNames.PostItem)]
    public sealed class PostItem : Item
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="PostItem"/>. To bind to an existing post item, use PostItem.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the e-mail message will be bound.</param>
        public PostItem(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PostItem"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal PostItem(ItemAttachment parentAttachment)
            : base(parentAttachment)
        {
        }

        /// <summary>
        /// Binds to an existing post item and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the post item.</param>
        /// <param name="id">The Id of the post item to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An PostItem instance representing the post item corresponding to the specified Id.</returns>
        public static new PostItem Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<PostItem>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing post item and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the post item.</param>
        /// <param name="id">The Id of the post item to bind to.</param>
        /// <returns>An PostItem instance representing the post item corresponding to the specified Id.</returns>
        public static new PostItem Bind(ExchangeService service, ItemId id)
        {
            return PostItem.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return PostItemSchema.Instance;
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
        /// Creates a post reply to this post item.
        /// </summary>
        /// <returns>A PostReply that can be modified and saved.</returns>
        public PostReply CreatePostReply()
        {
            this.ThrowIfThisIsNew();

            return new PostReply(this);
        }

        /// <summary>
        /// Posts a reply to this post item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">Body prefix.</param>
        public void PostReply(MessageBody bodyPrefix)
        {
            PostReply postReply = this.CreatePostReply();

            postReply.BodyPrefix = bodyPrefix;

            postReply.Save();
        }

        /// <summary>
        /// Creates a e-mail reply response to the post item.
        /// </summary>
        /// <param name="replyAll">Indicates whether the reply should go to everyone involved in the thread.</param>
        /// <returns>A ResponseMessage representing the e-mail reply response that can subsequently be modified and sent.</returns>
        public ResponseMessage CreateReply(bool replyAll)
        {
            this.ThrowIfThisIsNew();

            return new ResponseMessage(
                this,
                replyAll ? ResponseMessageType.ReplyAll : ResponseMessageType.Reply);
        }

        /// <summary>
        /// Replies to the post item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the post item.</param>
        /// <param name="replyAll">Indicates whether the reply should be sent to everyone involved in the thread.</param>
        public void Reply(MessageBody bodyPrefix, bool replyAll)
        {
            ResponseMessage responseMessage = this.CreateReply(replyAll);

            responseMessage.BodyPrefix = bodyPrefix;

            responseMessage.SendAndSaveCopy();
        }

        /// <summary>
        /// Creates a forward response to the post item.
        /// </summary>
        /// <returns>A ResponseMessage representing the forward response that can subsequently be modified and sent.</returns>
        public ResponseMessage CreateForward()
        {
            this.ThrowIfThisIsNew();

            return new ResponseMessage(this, ResponseMessageType.Forward);
        }

        /// <summary>
        /// Forwards the post item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the post item.</param>
        /// <param name="toRecipients">The recipients to forward the post item to.</param>
        public void Forward(MessageBody bodyPrefix, params EmailAddress[] toRecipients)
        {
            this.Forward(bodyPrefix, (IEnumerable<EmailAddress>)toRecipients);
        }

        /// <summary>
        /// Forwards the post item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the post item.</param>
        /// <param name="toRecipients">The recipients to forward the post item to.</param>
        public void Forward(MessageBody bodyPrefix, IEnumerable<EmailAddress> toRecipients)
        {
            ResponseMessage responseMessage = this.CreateForward();

            responseMessage.BodyPrefix = bodyPrefix;
            responseMessage.ToRecipients.AddRange(toRecipients);

            responseMessage.SendAndSaveCopy();
        }

        #region Properties

        /// <summary>
        /// Gets the conversation index of the post item.
        /// </summary>
        public byte[] ConversationIndex
        {
            get { return (byte[])this.PropertyBag[EmailMessageSchema.ConversationIndex]; }
        }

        /// <summary>
        /// Gets the conversation topic of the post item.
        /// </summary>
        public string ConversationTopic
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.ConversationTopic]; }
        }

        /// <summary>
        /// Gets or sets the "on behalf" poster of the post item.
        /// </summary>
        public EmailAddress From
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.From]; }
            set { this.PropertyBag[EmailMessageSchema.From] = value; }
        }

        /// <summary>
        /// Gets the Internet message Id of the post item.
        /// </summary>
        public string InternetMessageId
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.InternetMessageId]; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the post item is read.
        /// </summary>
        public bool IsRead
        {
            get { return (bool)this.PropertyBag[EmailMessageSchema.IsRead]; }
            set { this.PropertyBag[EmailMessageSchema.IsRead] = value; }
        }

        /// <summary>
        /// Gets the the date and time when the post item was posted.
        /// </summary>
        public DateTime PostedTime
        {
            get { return (DateTime)this.PropertyBag[PostItemSchema.PostedTime]; }
        }

        /// <summary>
        /// Gets or sets the references of the post item.
        /// </summary>
        public string References
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.References]; }
            set { this.PropertyBag[EmailMessageSchema.References] = value; }
        }

        /// <summary>
        /// Gets or sets the sender (poster) of the post item.
        /// </summary>
        public EmailAddress Sender
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.Sender]; }
            set { this.PropertyBag[EmailMessageSchema.Sender] = value; }
        }

        #endregion
    }
}