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
// <summary>Implements an attachment collection.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;

    /// <summary>
    /// Represents an item's attachment collection.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class AttachmentCollection : ComplexPropertyCollection<Attachment>, IOwnedProperty
    {
        #region Fields

        /// <summary>
        /// The item owner that owns this attachment collection
        /// </summary>
        private Item owner;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of AttachmentCollection.
        /// </summary>
        internal AttachmentCollection()
            : base()
        {
        }

        #endregion

        #region Properties

        #region IOwnedProperty Members

        /// <summary>
        /// The owner of this attachment collection.
        /// </summary>
        ServiceObject IOwnedProperty.Owner
        {
            get
            {
                return this.owner;
            }

            set
            {
                Item item = value as Item;

                EwsUtilities.Assert(
                    item != null,
                    "AttachmentCollection.IOwnedProperty.set_Owner",
                    "value is not a descendant of ItemBase");

                this.owner = item;
            }
        }

        #endregion

        #endregion

        #region Methods

        /// <summary>
        /// Adds a file attachment to the collection.
        /// </summary>
        /// <param name="fileName">The name of the file representing the content of the attachment.</param>
        /// <returns>A FileAttachment instance.</returns>
        public FileAttachment AddFileAttachment(string fileName)
        {
            return this.AddFileAttachment(Path.GetFileName(fileName), fileName);
        }

        /// <summary>
        /// Adds a file attachment to the collection.
        /// </summary>
        /// <param name="name">The display name of the new attachment.</param>
        /// <param name="fileName">The name of the file representing the content of the attachment.</param>
        /// <returns>A FileAttachment instance.</returns>
        public FileAttachment AddFileAttachment(string name, string fileName)
        {
            FileAttachment fileAttachment = new FileAttachment(this.owner);
            fileAttachment.Name = name;
            fileAttachment.FileName = fileName;

            this.InternalAdd(fileAttachment);

            return fileAttachment;
        }

        /// <summary>
        /// Adds a file attachment to the collection.
        /// </summary>
        /// <param name="name">The display name of the new attachment.</param>
        /// <param name="contentStream">The stream from which to read the content of the attachment.</param>
        /// <returns>A FileAttachment instance.</returns>
        public FileAttachment AddFileAttachment(string name, Stream contentStream)
        {
            FileAttachment fileAttachment = new FileAttachment(this.owner);
            fileAttachment.Name = name;
            fileAttachment.ContentStream = contentStream;

            this.InternalAdd(fileAttachment);

            return fileAttachment;
        }

        /// <summary>
        /// Adds a file attachment to the collection.
        /// </summary>
        /// <param name="name">The display name of the new attachment.</param>
        /// <param name="content">A byte arrays representing the content of the attachment.</param>
        /// <returns>A FileAttachment instance.</returns>
        public FileAttachment AddFileAttachment(string name, byte[] content)
        {
            FileAttachment fileAttachment = new FileAttachment(this.owner);
            fileAttachment.Name = name;
            fileAttachment.Content = content;

            this.InternalAdd(fileAttachment);

            return fileAttachment;
        }

        /// <summary>
        /// Adds an item attachment to the collection
        /// </summary>
        /// <typeparam name="TItem">The type of the item to attach.</typeparam>
        /// <returns>An ItemAttachment instance.</returns>
        public ItemAttachment<TItem> AddItemAttachment<TItem>()
            where TItem : Item
        {
            if (typeof(TItem).GetCustomAttributes(typeof(AttachableAttribute), false).Length == 0)
            {
                throw new InvalidOperationException(
                    string.Format(
                        "Items of type {0} are not supported as attachments.",
                        typeof(TItem).Name));
            }

            ItemAttachment<TItem> itemAttachment = new ItemAttachment<TItem>(this.owner);
            itemAttachment.Item = (TItem)EwsUtilities.CreateItemFromItemClass(itemAttachment, typeof(TItem), true);

            this.InternalAdd(itemAttachment);

            return itemAttachment;
        }

        /// <summary>
        /// Removes all attachments from this collection.
        /// </summary>
        public void Clear()
        {
            this.InternalClear();
        }

        /// <summary>
        /// Removes the attachment at the specified index.
        /// </summary>
        /// <param name="index">Index of the attachment to remove.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.Count)
            {
                throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
            }

            this.InternalRemoveAt(index);
        }

        /// <summary>
        /// Removes the specified attachment.
        /// </summary>
        /// <param name="attachment">The attachment to remove.</param>
        /// <returns>True if the attachment was successfully removed from the collection, false otherwise.</returns>
        public bool Remove(Attachment attachment)
        {
            EwsUtilities.ValidateParam(attachment, "attachment");

            return this.InternalRemove(attachment);
        }

        /// <summary>
        /// Instantiate the appropriate attachment type depending on the current XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to determine the type of attachment to create.</param>
        /// <returns>An Attachment instance.</returns>
        internal override Attachment CreateComplexProperty(string xmlElementName)
        {
            switch (xmlElementName)
            {
                case XmlElementNames.FileAttachment:
                    return new FileAttachment(this.owner);
                case XmlElementNames.ItemAttachment:
                    return new ItemAttachment(this.owner);
                default:
                    return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>An Attachment instance.</returns>
        internal override Attachment CreateDefaultComplexProperty()
        {
            throw new JsonDeserializationNotImplementedException();
        }

        /// <summary>
        /// Determines the name of the XML element associated with the complexProperty parameter.
        /// </summary>
        /// <param name="complexProperty">The attachment object for which to determine the XML element name with.</param>
        /// <returns>The XML element name associated with the complexProperty parameter.</returns>
        internal override string GetCollectionItemXmlElementName(Attachment complexProperty)
        {
            if (complexProperty is FileAttachment)
            {
                return XmlElementNames.FileAttachment;
            }
            else
            {
                return XmlElementNames.ItemAttachment;
            }
        }

        /// <summary>
        /// Saves this collection by creating new attachment and deleting removed ones.
        /// </summary>
        internal void Save()
        {
            List<Attachment> attachments = new List<Attachment>();

            // Retrieve a list of attachments that have to be deleted.
            foreach (Attachment attachment in this.RemovedItems)
            {
                if (!attachment.IsNew)
                {
                    attachments.Add(attachment);
                }
            }

            // If any, delete them by calling the DeleteAttachment web method.
            if (attachments.Count > 0)
            {
                this.InternalDeleteAttachments(attachments);
            }

            attachments.Clear();

            // Retrieve a list of attachments that have to be created.
            foreach (Attachment attachment in this)
            {
                if (attachment.IsNew)
                {
                    attachments.Add(attachment);
                }
            }

            // If there are any, create them by calling the CreateAttachment web method.
            if (attachments.Count > 0)
            {
                if (this.owner.IsAttachment)
                {
                    this.InternalCreateAttachments(this.owner.ParentAttachment.Id, attachments);
                }
                else
                {
                    this.InternalCreateAttachments(this.owner.Id.UniqueId, attachments);
                }
            }

            // Process all of the item attachments in this collection.
            foreach (Attachment attachment in this)
            {
                ItemAttachment itemAttachment = attachment as ItemAttachment;
                if (itemAttachment != null)
                {
                    // Make sure item was created/loaded before trying to create/delete sub-attachments
                    if (itemAttachment.Item != null)
                    {
                        // Create/delete any sub-attachments
                        itemAttachment.Item.Attachments.Save();

                        // Clear the item's change log
                        itemAttachment.Item.ClearChangeLog();
                    }
                }
            }

            base.ClearChangeLog();
        }

        /// <summary>
        /// Determines whether there are any unsaved attachment collection changes.
        /// </summary>
        /// <returns>True if attachment adds or deletes haven't been processed yet.</returns>
        internal bool HasUnprocessedChanges()
        {
            // Any new attachments?
            foreach (Attachment attachment in this)
            {
                if (attachment.IsNew)
                {
                    return true;
                }
            }

            // Any pending deletions?
            foreach (Attachment attachment in this.RemovedItems)
            {
                if (!attachment.IsNew)
                {
                    return true;
                }
            }

            // Recurse: process item attachments to check for new or deleted sub-attachments.
            foreach (ItemAttachment itemAttachment in this.OfType<ItemAttachment>())
            {
                if (itemAttachment.Item != null)
                {
                    if (itemAttachment.Item.Attachments.HasUnprocessedChanges())
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Disables the change log clearing mechanism. Attachment collections are saved separately
        /// from the items they belong to.
        /// </summary>
        internal override void ClearChangeLog()
        {
            // Do nothing
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal void Validate()
        {
            // Validate all added attachments
            bool contactPhotoFound = false;

            for (int attachmentIndex = 0; attachmentIndex < this.AddedItems.Count; attachmentIndex++)
            {
                Attachment attachment = this.AddedItems[attachmentIndex];
                if (attachment.IsNew)
                {
                    // At the server side, only the last attachment with IsContactPhoto is kept, all other IsContactPhoto
                    // attachments are removed. CreateAttachment will generate AttachmentId for each of such attachments (although
                    // only the last one is valid).
                    // 
                    // With E14 SP2 CreateItemWithAttachment, such request will only return 1 AttachmentId; but the client
                    // expects to see all, so let us prevent such "invalid" request in the first place. 
                    // 
                    // The IsNew check is to still let CreateAttachmentRequest allow multiple IsContactPhoto attachments.
                    // 
                    if (this.owner.IsNew && this.owner.Service.RequestedServerVersion >= ExchangeVersion.Exchange2010_SP2)
                    {
                        FileAttachment fileAttachment = attachment as FileAttachment;

                        if (fileAttachment != null && fileAttachment.IsContactPhoto)
                        {
                            if (contactPhotoFound)
                            {
                                throw new ServiceValidationException(Strings.MultipleContactPhotosInAttachment);
                            }

                            contactPhotoFound = true;
                        }
                    }

                    attachment.Validate(attachmentIndex);
                }
            }
        }

        /// <summary>
        /// Calls the DeleteAttachment web method to delete a list of attachments.
        /// </summary>
        /// <param name="attachments">The attachments to delete.</param>
        private void InternalDeleteAttachments(IEnumerable<Attachment> attachments)
        {
            ServiceResponseCollection<DeleteAttachmentResponse> responses = this.owner.Service.DeleteAttachments(attachments);

            foreach (DeleteAttachmentResponse response in responses)
            {
                // We remove all attachments that were successfully deleted from the change log. We should never
                // receive a warning from EWS, so we ignore them.
                if (response.Result != ServiceResult.Error)
                {
                    this.RemoveFromChangeLog(response.Attachment);
                }
            }

            // TODO : Should we throw for warnings as well?
            if (responses.OverallResult == ServiceResult.Error)
            {
                throw new DeleteAttachmentException(responses, Strings.AtLeastOneAttachmentCouldNotBeDeleted);
            }
        }

        /// <summary>
        /// Calls the CreateAttachment web method to create a list of attachments.
        /// </summary>
        /// <param name="parentItemId">The Id of the parent item of the new attachments.</param>
        /// <param name="attachments">The attachments to create.</param>
        private void InternalCreateAttachments(string parentItemId, IEnumerable<Attachment> attachments)
        {
            ServiceResponseCollection<CreateAttachmentResponse> responses = this.owner.Service.CreateAttachments(parentItemId, attachments);

            foreach (CreateAttachmentResponse response in responses)
            {
                // We remove all attachments that were successfully created from the change log. We should never
                // receive a warning from EWS, so we ignore them.
                if (response.Result != ServiceResult.Error)
                {
                    this.RemoveFromChangeLog(response.Attachment);
                }
            }

            // TODO : Should we throw for warnings as well?
            if (responses.OverallResult == ServiceResult.Error)
            {
                throw new CreateAttachmentException(responses, Strings.AttachmentCreationFailed);
            }
        }
        #endregion
    }
}