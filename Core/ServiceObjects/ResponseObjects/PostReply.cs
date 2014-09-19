// ---------------------------------------------------------------------------
// <copyright file="PostReply.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PostReply class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a reply to a post item.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.PostReplyItem, ReturnedByServer = false)]
    public sealed class PostReply : ServiceObject
    {
        private Item referenceItem;

        /// <summary>
        /// Initializes a new instance of the <see cref="PostReply"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal PostReply(Item referenceItem)
            : base(referenceItem.Service)
        {
            EwsUtilities.Assert(
                referenceItem != null,
                "PostReply.ctor",
                "referenceItem is null");

            referenceItem.ThrowIfThisIsNew();

            this.referenceItem = referenceItem;
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return PostReplySchema.Instance;
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
        /// Create a PostItem response.
        /// </summary>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <returns>Created PostItem.</returns>
        internal PostItem InternalCreate(FolderId parentFolderId, MessageDisposition? messageDisposition)
        {
            ((ItemId)this.PropertyBag[ResponseObjectSchema.ReferenceItemId]).Assign(this.referenceItem.Id);

            List<Item> items = this.Service.InternalCreateResponseObject(
                this,
                parentFolderId,
                messageDisposition);
            
            PostItem postItem = EwsUtilities.FindFirstItemOfType<PostItem>(items);

            // This should never happen. If it does, we have a bug.
            EwsUtilities.Assert(
                postItem != null,
                "PostReply.InternalCreate",
                "postItem is null. The CreateItem call did not return the expected PostItem.");

            return postItem;
        }

        /// <summary>
        /// Loads the specified set of properties on the object.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        internal override void InternalLoad(PropertySet propertySet)
        {
            throw new InvalidOperationException(Strings.LoadingThisObjectTypeNotSupported);
        }

        /// <summary>
        /// Deletes the object.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
        /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
        internal override void InternalDelete(
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences)
        {
            throw new InvalidOperationException(Strings.DeletingThisObjectTypeNotAuthorized);
        }

        /// <summary>
        /// Saves the post reply in the same folder as the original post item. Calling this method results in a call to EWS.
        /// </summary>
        /// <returns>A PostItem representing the posted reply.</returns>
        public PostItem Save()
        {
            return (PostItem)this.InternalCreate(null, null);
        }

        /// <summary>
        /// Saves the post reply in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save the post reply.</param>
        /// <returns>A PostItem representing the posted reply.</returns>
        public PostItem Save(FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return (PostItem)this.InternalCreate(destinationFolderId, null);
        }

        /// <summary>
        /// Saves the post reply in a specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save the post reply.</param>
        /// <returns>A PostItem representing the posted reply.</returns>
        public PostItem Save(WellKnownFolderName destinationFolderName)
        {
            return (PostItem)this.InternalCreate(new FolderId(destinationFolderName), null);
        }

        #region Properties

        /// <summary>
        /// Gets or sets the subject of the post reply.
        /// </summary>
        public string Subject
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.Subject]; }
            set { this.PropertyBag[EmailMessageSchema.Subject] = value; }
        }

        /// <summary>
        /// Gets or sets the body of the post reply.
        /// </summary>
        public MessageBody Body
        {
            get { return (MessageBody)this.PropertyBag[ItemSchema.Body]; }
            set { this.PropertyBag[ItemSchema.Body] = value; }
        }

        /// <summary>
        /// Gets or sets the body prefix that should be prepended to the original post item's body.
        /// </summary>
        public MessageBody BodyPrefix
        {
            get { return (MessageBody)this.PropertyBag[ResponseObjectSchema.BodyPrefix]; }
            set { this.PropertyBag[ResponseObjectSchema.BodyPrefix] = value; }
        }

        #endregion
    }
}
