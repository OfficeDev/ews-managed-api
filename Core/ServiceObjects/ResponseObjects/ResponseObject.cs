// ---------------------------------------------------------------------------
// <copyright file="ResponseObject.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseObject class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents the base class for all responses that can be sent.
    /// </summary>
    /// <typeparam name="TMessage">Type of message.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class ResponseObject<TMessage> : ServiceObject
        where TMessage : EmailMessage
    {
        private Item referenceItem;

        /// <summary>
        /// Initializes a new instance of the <see cref="ResponseObject&lt;TMessage&gt;"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal ResponseObject(Item referenceItem)
            : base(referenceItem.Service)
        {
            EwsUtilities.Assert(
                referenceItem != null,
                "ResponseObject.ctor",
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
            return ResponseObjectSchema.Instance;
        }

        /// <summary>
        /// Loads the specified set of properties on the object.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        internal override void InternalLoad(PropertySet propertySet)
        {
            throw new NotSupportedException();
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
            throw new NotSupportedException();
        }

        /// <summary>
        /// Create the response object.
        /// </summary>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <returns>The list of items returned by EWS.</returns>
        internal List<Item> InternalCreate(FolderId destinationFolderId, MessageDisposition messageDisposition)
        {
            ((ItemId)this.PropertyBag[ResponseObjectSchema.ReferenceItemId]).Assign(this.referenceItem.Id);

            return this.Service.InternalCreateResponseObject(
                this,
                destinationFolderId,
                messageDisposition);
        }

        /// <summary>
        /// Saves the response in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save the response.</param>
        /// <returns>A TMessage that represents the response.</returns>
        public TMessage Save(FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return this.InternalCreate(destinationFolderId, MessageDisposition.SaveOnly)[0] as TMessage;
        }

        /// <summary>
        /// Saves the response in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save the response.</param>
        /// <returns>A TMessage that represents the response.</returns>
        public TMessage Save(WellKnownFolderName destinationFolderName)
        {
            return this.InternalCreate(new FolderId(destinationFolderName), MessageDisposition.SaveOnly)[0] as TMessage;
        }

        /// <summary>
        /// Saves the response in the Drafts folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <returns>A TMessage that represents the response.</returns>
        public TMessage Save()
        {
            return this.InternalCreate(null, MessageDisposition.SaveOnly)[0] as TMessage;
        }

        /// <summary>
        /// Sends this response without saving a copy. Calling this method results in a call to EWS.
        /// </summary>
        public void Send()
        {
            this.InternalCreate(null, MessageDisposition.SendOnly);
        }

        /// <summary>
        /// Sends this response and saves a copy in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save the copy of the message.</param>
        public void SendAndSaveCopy(FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            this.InternalCreate(destinationFolderId, MessageDisposition.SendAndSaveCopy);
        }

        /// <summary>
        /// Sends this response and saves a copy in the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save the copy of the message.</param>
        public void SendAndSaveCopy(WellKnownFolderName destinationFolderName)
        {
            this.InternalCreate(new FolderId(destinationFolderName), MessageDisposition.SendAndSaveCopy);
        }

        /// <summary>
        /// Sends this response and saves a copy in the Sent Items folder. Calling this method results in a call to EWS.
        /// </summary>
        public void SendAndSaveCopy()
        {
            this.InternalCreate(
                null,
                MessageDisposition.SendAndSaveCopy);
        }

        #region Properties

        /// <summary>
        /// Gets or sets a value indicating whether read receipts will be requested from recipients of this response.
        /// </summary>
        public bool IsReadReceiptRequested
        {
            get { return (bool)this.PropertyBag[EmailMessageSchema.IsReadReceiptRequested]; }
            set { this.PropertyBag[EmailMessageSchema.IsReadReceiptRequested] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether delivery receipts should be sent to the sender.
        /// </summary>
        public bool IsDeliveryReceiptRequested
        {
            get { return (bool)this.PropertyBag[EmailMessageSchema.IsDeliveryReceiptRequested]; }
            set { this.PropertyBag[EmailMessageSchema.IsDeliveryReceiptRequested] = value; }
        }

        #endregion
    }
}
