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
    using System.Linq;

    /// <summary>
    /// Represents a generic item. Properties available on items are defined in the ItemSchema class.
    /// </summary>
    [Attachable]
    [ServiceObjectDefinition(XmlElementNames.Item)]
    public class Item : ServiceObject
    {
        private ItemAttachment parentAttachment;

        /// <summary>
        /// Initializes an unsaved local instance of <see cref="Item"/>. To bind to an existing item, use Item.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the item will be bound.</param>
        internal Item(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Item"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal Item(ItemAttachment parentAttachment)
            : this(parentAttachment.Service)
        {
            EwsUtilities.Assert(
                parentAttachment != null,
                "Item.ctor",
                "parentAttachment is null");

            this.parentAttachment = parentAttachment;
        }

        /// <summary>
        /// Binds to an existing item, whatever its actual type is, and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the item.</param>
        /// <param name="id">The Id of the item to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An Item instance representing the item corresponding to the specified Id.</returns>
        public static Item Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<Item>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing item, whatever its actual type is, and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the item.</param>
        /// <param name="id">The Id of the item to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An Item instance representing the item corresponding to the specified Id.</returns>
        public static async System.Threading.Tasks.Task<Item> BindAsync(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return await service.BindToItemAsync<Item>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing item, whatever its actual type is, and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the item.</param>
        /// <param name="id">The Id of the item to bind to.</param>
        /// <returns>An Item instance representing the item corresponding to the specified Id.</returns>
        public static Item Bind(ExchangeService service, ItemId id)
        {
            return Item.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an existing item, whatever its actual type is, and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the item.</param>
        /// <param name="id">The Id of the item to bind to.</param>
        /// <returns>An Item instance representing the item corresponding to the specified Id.</returns>
        public static async System.Threading.Tasks.Task<Item> BindAsync(ExchangeService service, ItemId id)
        {
            return await Item.BindAsync(
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
            return ItemSchema.Instance;
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
        /// Throws exception if this is attachment.
        /// </summary>
        internal void ThrowIfThisIsAttachment()
        {
            if (this.IsAttachment)
            {
                throw new InvalidOperationException(Strings.OperationDoesNotSupportAttachments);
            }
        }

        /// <summary>
        /// The property definition for the Id of this object.
        /// </summary>
        /// <returns>A PropertyDefinition instance.</returns>
        internal override PropertyDefinition GetIdPropertyDefinition()
        {
            return ItemSchema.Id;
        }

        /// <summary>
        /// Loads the specified set of properties on the object.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        internal override void InternalLoad(PropertySet propertySet)
        {
            this.ThrowIfThisIsNew();
            this.ThrowIfThisIsAttachment();

            this.Service.InternalLoadPropertiesForItems(
                new Item[] { this },
                propertySet,
                ServiceErrorHandling.ThrowOnError);
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
            this.InternalDelete(deleteMode, sendCancellationsMode, affectedTaskOccurrences, false);
        }

        /// <summary>
        /// Deletes the object.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
        /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        internal void InternalDelete(
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            bool suppressReadReceipts)
        {
            this.ThrowIfThisIsNew();
            this.ThrowIfThisIsAttachment();

            // If sendCancellationsMode is null, use the default value that's appropriate for item type.
            if (!sendCancellationsMode.HasValue)
            {
                sendCancellationsMode = this.DefaultSendCancellationsMode;
            }

            // If affectedTaskOccurrences is null, use the default value that's appropriate for item type.
            if (!affectedTaskOccurrences.HasValue)
            {
                affectedTaskOccurrences = this.DefaultAffectedTaskOccurrences;
            }

            this.Service.DeleteItem(
                this.Id,
                deleteMode,
                sendCancellationsMode,
                affectedTaskOccurrences,
                suppressReadReceipts);
        }

        /// <summary>
        /// Create item.
        /// </summary>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <param name="sendInvitationsMode">The send invitations mode.</param>
        internal void InternalCreate(
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode)
        {
            this.ThrowIfThisIsNotNew();
            this.ThrowIfThisIsAttachment();

            if (this.IsNew || this.IsDirty)
            {
                this.Service.CreateItem(
                    this,
                    parentFolderId,
                    messageDisposition,
                    sendInvitationsMode.HasValue ? sendInvitationsMode : this.DefaultSendInvitationsMode);

                this.Attachments.Save();
            }
        }

        /// <summary>
        /// Update item.
        /// </summary>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <param name="sendInvitationsOrCancellationsMode">The send invitations or cancellations mode.</param>
        /// <returns>Updated item.</returns>
        internal Item InternalUpdate(
            FolderId parentFolderId,
            ConflictResolutionMode conflictResolutionMode,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode)
        {
            return this.InternalUpdate(parentFolderId, conflictResolutionMode, messageDisposition, sendInvitationsOrCancellationsMode, false);
        }

        /// <summary>
        /// Update item.
        /// </summary>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <param name="sendInvitationsOrCancellationsMode">The send invitations or cancellations mode.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>Updated item.</returns>
        internal Item InternalUpdate(
            FolderId parentFolderId,
            ConflictResolutionMode conflictResolutionMode,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            bool suppressReadReceipts)
        {
            this.ThrowIfThisIsNew();
            this.ThrowIfThisIsAttachment();

            Item returnedItem = null;

            if (this.IsDirty && this.PropertyBag.GetIsUpdateCallNecessary())
            {
                returnedItem = this.Service.UpdateItem(
                    this,
                    parentFolderId,
                    conflictResolutionMode,
                    messageDisposition,
                    sendInvitationsOrCancellationsMode.HasValue ? sendInvitationsOrCancellationsMode : this.DefaultSendInvitationsOrCancellationsMode,
                    suppressReadReceipts);
            }

            // Regardless of whether item is dirty or not, if it has unprocessed
            // attachment changes, validate them and process now.
            if (this.HasUnprocessedAttachmentChanges())
            {
                this.Attachments.Validate();
                this.Attachments.Save();
            }

            return returnedItem;
        }

        /// <summary>
        /// Gets a value indicating whether this instance has unprocessed attachment collection changes.
        /// </summary>
        internal bool HasUnprocessedAttachmentChanges()
        {
            return this.Attachments.HasUnprocessedChanges();
        }

        /// <summary>
        /// Gets the parent attachment of this item.
        /// </summary>
        internal ItemAttachment ParentAttachment
        {
            get { return this.parentAttachment; }
        }

        /// <summary>
        /// Gets Id of the root item for this item.
        /// </summary>
        internal ItemId RootItemId
        {
            get
            {
                if (this.IsAttachment && this.ParentAttachment.Owner != null)
                {
                    return this.ParentAttachment.Owner.RootItemId;
                }
                else
                {
                    return this.Id;
                }
            }
        }

        /// <summary>
        /// Deletes the item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        public void Delete(DeleteMode deleteMode)
        {
            this.Delete(deleteMode, false);
        }

        /// <summary>
        /// Deletes the item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        public void Delete(DeleteMode deleteMode, bool suppressReadReceipts)
        {
            this.InternalDelete(deleteMode, null, null, suppressReadReceipts);
        }

        /// <summary>
        /// Saves this item in a specific folder. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to save this item.</param>
        public void Save(FolderId parentFolderId)
        {
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");

            this.InternalCreate(
                parentFolderId,
                MessageDisposition.SaveOnly,
                null);
        }

        /// <summary>
        /// Saves this item in a specific folder. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to save this item.</param>
        public void Save(WellKnownFolderName parentFolderName)
        {
            this.InternalCreate(
                new FolderId(parentFolderName),
                MessageDisposition.SaveOnly,
                null);
        }

        /// <summary>
        /// Saves this item in the default folder based on the item's type (for example, an e-mail message is saved to the Drafts folder).
        /// Calling this method results in at least one call to EWS. Mutliple calls to EWS might be made if attachments have been added.
        /// </summary>
        public void Save()
        {
            this.InternalCreate(
                null,
                MessageDisposition.SaveOnly,
                null);
        }

        /// <summary>
        /// Applies the local changes that have been made to this item. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added or removed.
        /// </summary>
        /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
        public void Update(ConflictResolutionMode conflictResolutionMode)
        {
            this.Update(conflictResolutionMode, false);
        }

        /// <summary>
        /// Applies the local changes that have been made to this item. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added or removed.
        /// </summary>
        /// <param name="conflictResolutionMode">The conflict resolution mode.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        public void Update(ConflictResolutionMode conflictResolutionMode, bool suppressReadReceipts)
        {
            this.InternalUpdate(
                null /* parentFolder */,
                conflictResolutionMode,
                MessageDisposition.SaveOnly,
                null,
                suppressReadReceipts);
        }

        /// <summary>
        /// Creates a copy of this item in the specified folder. Calling this method results in a call to EWS.
        /// <para>
        /// Copy returns null if the copy operation is across two mailboxes or between a mailbox and a
        /// public folder.
        /// </para>
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to create a copy of this item.</param>
        /// <returns>The copy of this item.</returns>
        public Item Copy(FolderId destinationFolderId)
        {
            this.ThrowIfThisIsNew();
            this.ThrowIfThisIsAttachment();

            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return this.Service.CopyItem(this.Id, destinationFolderId);
        }

        /// <summary>
        /// Creates a copy of this item in the specified folder. Calling this method results in a call to EWS.
        /// <para>
        /// Copy returns null if the copy operation is across two mailboxes or between a mailbox and a
        /// public folder.
        /// </para>
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to create a copy of this item.</param>
        /// <returns>The copy of this item.</returns>
        public Item Copy(WellKnownFolderName destinationFolderName)
        {
            return this.Copy(new FolderId(destinationFolderName));
        }

        /// <summary>
        /// Moves this item to a the specified folder. Calling this method results in a call to EWS.
        /// <para>
        /// Move returns null if the move operation is across two mailboxes or between a mailbox and a
        /// public folder.
        /// </para>
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder to which to move this item.</param>
        /// <returns>The moved copy of this item.</returns>
        public Item Move(FolderId destinationFolderId)
        {
            this.ThrowIfThisIsNew();
            this.ThrowIfThisIsAttachment();

            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return this.Service.MoveItem(this.Id, destinationFolderId);
        }

        /// <summary>
        /// Moves this item to a the specified folder. Calling this method results in a call to EWS.
        /// <para>
        /// Move returns null if the move operation is across two mailboxes or between a mailbox and a
        /// public folder.
        /// </para>
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder to which to move this item.</param>
        /// <returns>The moved copy of this item.</returns>
        public Item Move(WellKnownFolderName destinationFolderName)
        {
            return this.Move(new FolderId(destinationFolderName));
        }

        /// <summary>
        /// Sets the extended property.
        /// </summary>
        /// <param name="extendedPropertyDefinition">The extended property definition.</param>
        /// <param name="value">The value.</param>
        public void SetExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition, object value)
        {
            this.ExtendedProperties.SetExtendedProperty(extendedPropertyDefinition, value);
        }

        /// <summary>
        /// Removes an extended property.
        /// </summary>
        /// <param name="extendedPropertyDefinition">The extended property definition.</param>
        /// <returns>True if property was removed.</returns>
        public bool RemoveExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition)
        {
            return this.ExtendedProperties.RemoveExtendedProperty(extendedPropertyDefinition);
        }

        /// <summary>
        /// Gets a list of extended properties defined on this object.
        /// </summary>
        /// <returns>Extended properties collection.</returns>
        internal override ExtendedPropertyCollection GetExtendedProperties()
        {
            return this.ExtendedProperties;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            this.Attachments.Validate();

            // Flag parameter is only valid for Exchange2013 or higher
            //
            Flag flag;
            if (this.TryGetProperty<Flag>(ItemSchema.Flag, out flag) && flag != null)
            {
                if (this.Service.RequestedServerVersion < ExchangeVersion.Exchange2013)
                {
                    throw new ServiceVersionException(
                        string.Format(
                            Strings.ParameterIncompatibleWithRequestVersion,
                            "Flag",
                            ExchangeVersion.Exchange2013));
                }

                flag.Validate();
            }
        }

        /// <summary>
        /// Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
        /// or UpdateItem request so this item can be property saved or updated.
        /// </summary>
        /// <param name="isUpdateOperation">Indicates whether the operation being petrformed is an update operation.</param>
        /// <returns>
        ///     <c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.
        /// </returns>
        internal override bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
        {
            // Starting E14SP2, attachment will be sent along with CreateItem requests. 
            // if the attachment used to require the Timezone header, CreateItem request should do so too.
            //
            if (!isUpdateOperation &&
                (this.Service.RequestedServerVersion >= ExchangeVersion.Exchange2010_SP2))
            {
                foreach (ItemAttachment itemAttachment in this.Attachments.OfType<ItemAttachment>())
                {
                    if ((itemAttachment.Item != null) && itemAttachment.Item.GetIsTimeZoneHeaderRequired(false /* isUpdateOperation */))
                    {
                        return true;
                    }
                }
            }

            return base.GetIsTimeZoneHeaderRequired(isUpdateOperation);
        }

        #region Properties

        /// <summary>
        /// Gets a value indicating whether the item is an attachment.
        /// </summary>
        public bool IsAttachment
        {
            get { return this.parentAttachment != null; }
        }

        /// <summary>
        /// Gets a value indicating whether this object is a real store item, or if it's a local object
        /// that has yet to be saved. 
        /// </summary>
        public override bool IsNew
        {
            get
            {
                // Item attachments don't have an Id, need to check whether the
                // parentAttachment is new or not.
                if (this.IsAttachment)
                {
                    return this.ParentAttachment.IsNew;
                }
                else
                {
                    return base.IsNew;
                }
            }
        }

        /// <summary>
        /// Gets the Id of this item.
        /// </summary>
        public ItemId Id
        {
            get { return (ItemId)this.PropertyBag[this.GetIdPropertyDefinition()]; }
        }

        /// <summary>
        /// Get or sets the MIME content of this item.
        /// </summary>
        public MimeContent MimeContent
        {
            get { return (MimeContent)this.PropertyBag[ItemSchema.MimeContent]; }
            set { this.PropertyBag[ItemSchema.MimeContent] = value; }
        }

        /// <summary>
        /// Get or sets the MimeContentUTF8 of this item.
        /// </summary>
        public MimeContentUTF8 MimeContentUTF8
        {
            get { return (MimeContentUTF8)this.PropertyBag[ItemSchema.MimeContentUTF8]; }
            set { this.PropertyBag[ItemSchema.MimeContentUTF8] = value; }
        }

        /// <summary>
        /// Gets the Id of the parent folder of this item.
        /// </summary>
        public FolderId ParentFolderId
        {
            get { return (FolderId)this.PropertyBag[ItemSchema.ParentFolderId]; }
        }

        /// <summary>
        /// Gets or sets the sensitivity of this item.
        /// </summary>
        public Sensitivity Sensitivity
        {
            get { return (Sensitivity)this.PropertyBag[ItemSchema.Sensitivity]; }
            set { this.PropertyBag[ItemSchema.Sensitivity] = value; }
        }

        /// <summary>
        /// Gets a list of the attachments to this item.
        /// </summary>
        public AttachmentCollection Attachments
        {
            get { return (AttachmentCollection)this.PropertyBag[ItemSchema.Attachments]; }
        }

        /// <summary>
        /// Gets the time when this item was received.
        /// </summary>
        public DateTime DateTimeReceived
        {
            get { return (DateTime)this.PropertyBag[ItemSchema.DateTimeReceived]; }
        }

        /// <summary>
        /// Gets the size of this item.
        /// </summary>
        public int Size
        {
            get { return (int)this.PropertyBag[ItemSchema.Size]; }
        }

        /// <summary>
        /// Gets or sets the list of categories associated with this item.
        /// </summary>
        public StringList Categories
        {
            get { return (StringList)this.PropertyBag[ItemSchema.Categories]; }
            set { this.PropertyBag[ItemSchema.Categories] = value; }
        }

        /// <summary>
        /// Gets or sets the culture associated with this item.
        /// </summary>
        public string Culture
        {
            get { return (string)this.PropertyBag[ItemSchema.Culture]; }
            set { this.PropertyBag[ItemSchema.Culture] = value; }
        }

        /// <summary>
        /// Gets or sets the importance of this item.
        /// </summary>
        public Importance Importance
        {
            get { return (Importance)this.PropertyBag[ItemSchema.Importance]; }
            set { this.PropertyBag[ItemSchema.Importance] = value; }
        }

        /// <summary>
        /// Gets or sets the In-Reply-To reference of this item.
        /// </summary>
        public string InReplyTo
        {
            get { return (string)this.PropertyBag[ItemSchema.InReplyTo]; }
            set { this.PropertyBag[ItemSchema.InReplyTo] = value; }
        }

        /// <summary>
        /// Gets a value indicating whether the message has been submitted to be sent.
        /// </summary>
        public bool IsSubmitted
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsSubmitted]; }
        }

        /// <summary>
        /// Gets a value indicating whether this is an associated item.
        /// </summary>
        public bool IsAssociated
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsAssociated]; }
        }

        /// <summary>
        /// Gets a value indicating whether the item is is a draft. An item is a draft when it has not yet been sent.
        /// </summary>
        public bool IsDraft
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsDraft]; }
        }

        /// <summary>
        /// Gets a value indicating whether the item has been sent by the current authenticated user.
        /// </summary>
        public bool IsFromMe
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsFromMe]; }
        }

        /// <summary>
        /// Gets a value indicating whether the item is a resend of another item.
        /// </summary>
        public bool IsResend
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsResend]; }
        }

        /// <summary>
        /// Gets a value indicating whether the item has been modified since it was created.
        /// </summary>
        public bool IsUnmodified
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsUnmodified]; }
        }

        /// <summary>
        /// Gets a list of Internet headers for this item.
        /// </summary>
        public InternetMessageHeaderCollection InternetMessageHeaders
        {
            get { return (InternetMessageHeaderCollection)this.PropertyBag[ItemSchema.InternetMessageHeaders]; }
        }

        /// <summary>
        /// Gets the date and time this item was sent.
        /// </summary>
        public DateTime DateTimeSent
        {
            get { return (DateTime)this.PropertyBag[ItemSchema.DateTimeSent]; }
        }

        /// <summary>
        /// Gets the date and time this item was created.
        /// </summary>
        public DateTime DateTimeCreated
        {
            get { return (DateTime)this.PropertyBag[ItemSchema.DateTimeCreated]; }
        }

        /// <summary>
        /// Gets a value indicating which response actions are allowed on this item. Examples of response actions are Reply and Forward.
        /// </summary>
        public ResponseActions AllowedResponseActions
        {
            get { return (ResponseActions)this.PropertyBag[ItemSchema.AllowedResponseActions]; }
        }

        /// <summary>
        /// Gets or sets the date and time when the reminder is due for this item.
        /// </summary>
        public DateTime ReminderDueBy
        {
            get { return (DateTime)this.PropertyBag[ItemSchema.ReminderDueBy]; }
            set { this.PropertyBag[ItemSchema.ReminderDueBy] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a reminder is set for this item.
        /// </summary>
        public bool IsReminderSet
        {
            get { return (bool)this.PropertyBag[ItemSchema.IsReminderSet]; }
            set { this.PropertyBag[ItemSchema.IsReminderSet] = value; }
        }

        /// <summary>
        /// Gets or sets the number of minutes before the start of this item when the reminder should be triggered.
        /// </summary>
        public int ReminderMinutesBeforeStart
        {
            get { return (int)this.PropertyBag[ItemSchema.ReminderMinutesBeforeStart]; }
            set { this.PropertyBag[ItemSchema.ReminderMinutesBeforeStart] = value; }
        }

        /// <summary>
        /// Gets a text summarizing the Cc receipients of this item.
        /// </summary>
        public string DisplayCc
        {
            get { return (string)this.PropertyBag[ItemSchema.DisplayCc]; }
        }

        /// <summary>
        /// Gets a text summarizing the To recipients of this item.
        /// </summary>
        public string DisplayTo
        {
            get { return (string)this.PropertyBag[ItemSchema.DisplayTo]; }
        }

        /// <summary>
        /// Gets a value indicating whether the item has attachments.
        /// </summary>
        public bool HasAttachments
        {
            get { return (bool)this.PropertyBag[ItemSchema.HasAttachments]; }
        }

        /// <summary>
        /// Gets or sets the body of this item.
        /// </summary>
        public MessageBody Body
        {
            get { return (MessageBody)this.PropertyBag[ItemSchema.Body]; }
            set { this.PropertyBag[ItemSchema.Body] = value; }
        }

        /// <summary>
        /// Gets or sets the custom class name of this item.
        /// </summary>
        public string ItemClass
        {
            get { return (string)this.PropertyBag[ItemSchema.ItemClass]; }
            set { this.PropertyBag[ItemSchema.ItemClass] = value; }
        }

        /// <summary>
        /// Gets or sets the subject of this item.
        /// </summary>
        public string Subject
        {
            get { return (string)this.PropertyBag[ItemSchema.Subject]; }
            set { this.SetSubject(value); }
        }

        /// <summary>
        /// Gets the query string that should be appended to the Exchange Web client URL to open this item using the appropriate read form in a web browser.
        /// </summary>
        public string WebClientReadFormQueryString
        {
            get { return (string)this.PropertyBag[ItemSchema.WebClientReadFormQueryString]; }
        }

        /// <summary>
        /// Gets the query string that should be appended to the Exchange Web client URL to open this item using the appropriate edit form in a web browser.
        /// </summary>
        public string WebClientEditFormQueryString
        {
            get { return (string)this.PropertyBag[ItemSchema.WebClientEditFormQueryString]; }
        }

        /// <summary>
        /// Gets a list of extended properties defined on this item.
        /// </summary>
        public ExtendedPropertyCollection ExtendedProperties
        {
            get { return (ExtendedPropertyCollection)this.PropertyBag[ServiceObjectSchema.ExtendedProperties]; }
        }

        /// <summary>
        /// Gets a value indicating the effective rights the current authenticated user has on this item.
        /// </summary>
        public EffectiveRights EffectiveRights
        {
            get { return (EffectiveRights)this.PropertyBag[ItemSchema.EffectiveRights]; }
        }

        /// <summary>
        /// Gets the name of the user who last modified this item.
        /// </summary>
        public string LastModifiedName
        {
            get { return (string)this.PropertyBag[ItemSchema.LastModifiedName]; }
        }

        /// <summary>
        /// Gets the date and time this item was last modified.
        /// </summary>
        public DateTime LastModifiedTime
        {
            get { return (DateTime)this.PropertyBag[ItemSchema.LastModifiedTime]; }
        }

        /// <summary>
        /// Gets the Id of the conversation this item is part of.
        /// </summary>
        public ConversationId ConversationId
        {
            get { return (ConversationId)this.PropertyBag[ItemSchema.ConversationId]; }
        }

        /// <summary>
        /// Gets the body part that is unique to the conversation this item is part of.
        /// </summary>
        public UniqueBody UniqueBody
        {
            get { return (UniqueBody)this.PropertyBag[ItemSchema.UniqueBody]; }
        }

        /// <summary>
        /// Gets the store entry id.
        /// </summary>
        public byte[] StoreEntryId
        {
            get { return (byte[])this.PropertyBag[ItemSchema.StoreEntryId]; }
        }

        /// <summary>
        /// Gets the item instance key.
        /// </summary>
        public byte[] InstanceKey
        {
            get { return (byte[])this.PropertyBag[ItemSchema.InstanceKey]; }
        }

        /// <summary>
        /// Get or set the Flag value for this item.
        /// </summary>
        public Flag Flag
        {
            get { return (Flag)this.PropertyBag[ItemSchema.Flag]; }
            set { this.PropertyBag[ItemSchema.Flag] = value; }
        }

        /// <summary>
        /// Gets the normalized body of the item.
        /// </summary>
        public NormalizedBody NormalizedBody
        {
            get { return (NormalizedBody)this.PropertyBag[ItemSchema.NormalizedBody]; }
        }

        /// <summary>
        /// Gets the EntityExtractionResult of the item.
        /// </summary>
        public EntityExtractionResult EntityExtractionResult
        {
            get { return (EntityExtractionResult)this.PropertyBag[ItemSchema.EntityExtractionResult]; }
        }

        /// <summary>
        /// Gets or sets the policy tag.
        /// </summary>
        public PolicyTag PolicyTag
        {
            get { return (PolicyTag)this.PropertyBag[ItemSchema.PolicyTag]; }
            set { this.PropertyBag[ItemSchema.PolicyTag] = value; }
        }

        /// <summary>
        /// Gets or sets the archive tag.
        /// </summary>
        public ArchiveTag ArchiveTag
        {
            get { return (ArchiveTag)this.PropertyBag[ItemSchema.ArchiveTag]; }
            set { this.PropertyBag[ItemSchema.ArchiveTag] = value; }
        }

        /// <summary>
        /// Gets the retention date.
        /// </summary>
        public DateTime? RetentionDate
        {
            get { return (DateTime?)this.PropertyBag[ItemSchema.RetentionDate]; }
        }

        /// <summary>
        /// Gets the item Preview.
        /// </summary>
        public string Preview
        {
            get { return (string)this.PropertyBag[ItemSchema.Preview]; }
        }

        /// <summary>
        /// Gets the text body of the item.
        /// </summary>
        public TextBody TextBody
        {
            get { return (TextBody)this.PropertyBag[ItemSchema.TextBody]; }
        }

        /// <summary>
        /// Gets the icon index.
        /// </summary>
        public IconIndex IconIndex
        {
            get { return (IconIndex)this.PropertyBag[ItemSchema.IconIndex]; }
        }

        /// <summary>
        /// Gets the default setting for how to treat affected task occurrences on Delete.
        /// Subclasses will override this for different default behavior.
        /// </summary>
        internal virtual AffectedTaskOccurrence? DefaultAffectedTaskOccurrences
        {
            get { return null; }
        }

        /// <summary>
        /// Gets the default setting for sending cancellations on Delete.
        /// Subclasses will override this for different default behavior.
        /// </summary>
        internal virtual SendCancellationsMode? DefaultSendCancellationsMode
        {
            get { return null; }
        }

        /// <summary>
        /// Gets the default settings for sending invitations on Save.
        /// Subclasses will override this for different default behavior.
        /// </summary>
        internal virtual SendInvitationsMode? DefaultSendInvitationsMode
        {
            get { return null; }
        }

        /// <summary>
        /// Gets the default settings for sending invitations or cancellations on Update.
        /// Subclasses will override this for different default behavior.
        /// </summary>
        internal virtual SendInvitationsOrCancellationsMode? DefaultSendInvitationsOrCancellationsMode
        {
            get { return null; }
        }

        /// <summary>
        /// Sets the subject.
        /// </summary>
        /// <param name="subject">The subject.</param>
        internal virtual void SetSubject(string subject)
        {
            this.PropertyBag[ItemSchema.Subject] = subject;
        }

        #endregion
    }
}