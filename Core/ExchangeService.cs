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

using System.Security.Cryptography;

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Autodiscover;
    using Microsoft.Exchange.WebServices.Data.Enumerations;
    using Microsoft.Exchange.WebServices.Data.Groups;

    /// <summary>
    /// Represents a binding to the Exchange Web Services.
    /// </summary>
    public sealed class ExchangeService : ExchangeServiceBase
    {
        #region Constants

        private const string TargetServerVersionHeaderName = "X-EWS-TargetVersion";

        #endregion

        #region Fields

        private Uri url;
        private CultureInfo preferredCulture;
        private DateTimePrecision dateTimePrecision = DateTimePrecision.Default;
        private ImpersonatedUserId impersonatedUserId;
        private PrivilegedUserId privilegedUserId;
        private ManagementRoles managementRoles;
        private IFileAttachmentContentHandler fileAttachmentContentHandler;
        private UnifiedMessaging unifiedMessaging;
        private bool enableScpLookup = true;
        private bool traceEnablePrettyPrinting = true;
        private string targetServerVersion = null;

        #endregion

        #region Response object operations

        /// <summary>
        /// Create response object.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <returns>The list of items created or modified as a result of the "creation" of the response object.</returns>
        internal List<Item> InternalCreateResponseObject(
            ServiceObject responseObject,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition)
        {
            CreateResponseObjectRequest request = new CreateResponseObjectRequest(this, ServiceErrorHandling.ThrowOnError);

            request.ParentFolderId = parentFolderId;
            request.Items = new ServiceObject[] { responseObject };
            request.MessageDisposition = messageDisposition;

            ServiceResponseCollection<CreateResponseObjectResponse> responses = request.Execute();

            return responses[0].Items;
        }

        #endregion

        #region Folder operations

        /// <summary>
        /// Creates a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="parentFolderId">The parent folder id.</param>
        internal void CreateFolder(
            Folder folder,
            FolderId parentFolderId)
        {
            CreateFolderRequest request = new CreateFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.Folders = new Folder[] { folder };
            request.ParentFolderId = parentFolderId;

            request.Execute();
        }

        /// <summary>
        /// Updates a folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        internal void UpdateFolder(Folder folder)
        {
            UpdateFolderRequest request = new UpdateFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.Folders.Add(folder);

            request.Execute();
        }

        /// <summary>
        /// Copies a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <returns>Copy of folder.</returns>
        internal Folder CopyFolder(
            FolderId folderId,
            FolderId destinationFolderId)
        {
            CopyFolderRequest request = new CopyFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.DestinationFolderId = destinationFolderId;
            request.FolderIds.Add(folderId);

            ServiceResponseCollection<MoveCopyFolderResponse> responses = request.Execute();

            return responses[0].Folder;
        }

        /// <summary>
        /// Move a folder.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <returns>Moved folder.</returns>
        internal Folder MoveFolder(
            FolderId folderId,
            FolderId destinationFolderId)
        {
            MoveFolderRequest request = new MoveFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.DestinationFolderId = destinationFolderId;
            request.FolderIds.Add(folderId);

            ServiceResponseCollection<MoveCopyFolderResponse> responses = request.Execute();

            return responses[0].Folder;
        }

        /// <summary>
        /// Finds folders.
        /// </summary>
        /// <param name="parentFolderIds">The parent folder ids.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <param name="errorHandlingMode">Indicates the type of error handling should be done.</param>
        /// <returns>Collection of service responses.</returns>
        private ServiceResponseCollection<FindFolderResponse> InternalFindFolders(
            IEnumerable<FolderId> parentFolderIds,
            SearchFilter searchFilter,
            FolderView view,
            ServiceErrorHandling errorHandlingMode)
        {
            FindFolderRequest request = new FindFolderRequest(this, errorHandlingMode);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.SearchFilter = searchFilter;
            request.View = view;

            return request.Execute();
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for folders.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindFoldersResults FindFolders(FolderId parentFolderId, SearchFilter searchFilter, FolderView view)
        {
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindFolderResponse> responses = this.InternalFindFolders(
                new FolderId[] { parentFolderId },
                searchFilter,
                view,
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Results;
        }
        
        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of each of the specified folders.
        /// </summary>
        /// <param name="parentFolderIds">The Ids of the folders in which to search for folders.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public ServiceResponseCollection<FindFolderResponse> FindFolders(IEnumerable<FolderId> parentFolderIds, SearchFilter searchFilter, FolderView view)
        {
            EwsUtilities.ValidateParam(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            return this.InternalFindFolders(
                parentFolderIds,
                searchFilter,
                view,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for folders.</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindFoldersResults FindFolders(FolderId parentFolderId, FolderView view)
        {
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");
            EwsUtilities.ValidateParam(view, "view");

            ServiceResponseCollection<FindFolderResponse> responses = this.InternalFindFolders(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                view,
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for folders.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindFoldersResults FindFolders(WellKnownFolderName parentFolderName, SearchFilter searchFilter, FolderView view)
        {
            return this.FindFolders(new FolderId(parentFolderName), searchFilter, view);
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for folders.</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindFoldersResults FindFolders(WellKnownFolderName parentFolderName, FolderView view)
        {
            return this.FindFolders(new FolderId(parentFolderName), view);
        }

        /// <summary>
        /// Load specified properties for a folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="propertySet">The property set.</param>
        internal void LoadPropertiesForFolder(
            Folder folder,
            PropertySet propertySet)
        {
            EwsUtilities.ValidateParam(folder, "folder");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            GetFolderRequestForLoad request = new GetFolderRequestForLoad(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folder);
            request.PropertySet = propertySet;

            request.Execute();
        }

        /// <summary>
        /// Binds to a folder.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Folder</returns>
        internal Folder BindToFolder(FolderId folderId, PropertySet propertySet)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            ServiceResponseCollection<GetFolderResponse> responses = this.InternalBindToFolders(
                new[] { folderId },
                propertySet,
                ServiceErrorHandling.ThrowOnError
            );

            return responses[0].Folder;
        }

        /// <summary>
        /// Binds to folder.
        /// </summary>
        /// <typeparam name="TFolder">The type of the folder.</typeparam>
        /// <param name="folderId">The folder id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Folder</returns>
        internal TFolder BindToFolder<TFolder>(FolderId folderId, PropertySet propertySet)
            where TFolder : Folder
        {
            Folder result = this.BindToFolder(folderId, propertySet);

            if (result is TFolder)
            {
                return (TFolder)result;
            }
            else
            {
                throw new ServiceLocalException(
                    string.Format(
                        Strings.FolderTypeNotCompatible,
                        result.GetType().Name,
                        typeof(TFolder).Name));
            }
        }

        /// <summary>
        /// Binds to multiple folders in a single call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folders to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified folder Ids.</returns>
        public ServiceResponseCollection<GetFolderResponse> BindToFolders(
            IEnumerable<FolderId> folderIds,
            PropertySet propertySet)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            return this.InternalBindToFolders(
                folderIds,
                propertySet,
                ServiceErrorHandling.ReturnErrors
            );
        }

        /// <summary>
        /// Binds to multiple folders in a single call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folders to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified folder Ids.</returns>
        private ServiceResponseCollection<GetFolderResponse> InternalBindToFolders(
            IEnumerable<FolderId> folderIds,
            PropertySet propertySet,
            ServiceErrorHandling errorHandling)
        {
            GetFolderRequest request = new GetFolderRequest(this, errorHandling);

            request.FolderIds.AddRange(folderIds);
            request.PropertySet = propertySet;

            return request.Execute();
        }

        /// <summary>
        /// Deletes a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="deleteMode">The delete mode.</param>
        internal void DeleteFolder(
            FolderId folderId,
            DeleteMode deleteMode)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");

            DeleteFolderRequest request = new DeleteFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folderId);
            request.DeleteMode = deleteMode;

            request.Execute();
        }

        /// <summary>
        /// Empties a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="deleteMode">The delete mode.</param>
        /// <param name="deleteSubFolders">if set to <c>true</c> empty folder should also delete sub folders.</param>
        internal void EmptyFolder(
            FolderId folderId,
            DeleteMode deleteMode,
            bool deleteSubFolders)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");

            EmptyFolderRequest request = new EmptyFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folderId);
            request.DeleteMode = deleteMode;
            request.DeleteSubFolders = deleteSubFolders;

            request.Execute();
        }

        /// <summary>
        /// Marks all items in folder as read/unread. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="readFlag">If true, items marked as read, otherwise unread.</param>
        /// <param name="suppressReadReceipts">If true, suppress read receipts for items.</param>
        internal void MarkAllItemsAsRead(
            FolderId folderId,
            bool readFlag,
            bool suppressReadReceipts)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "MarkAllItemsAsRead");

            MarkAllItemsAsReadRequest request = new MarkAllItemsAsReadRequest(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folderId);
            request.ReadFlag = readFlag;
            request.SuppressReadReceipts = suppressReadReceipts;

            request.Execute();
        }

        #endregion

        #region Item operations

        /// <summary>
        /// Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact, PostItem, Task and Item.
        /// CreateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to create.</param>
        /// <param name="parentFolderId">The Id of the folder in which to place the newly created items. If null, items are created in their default folders.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsMode">Indicates if and how invitations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>A ServiceResponseCollection providing creation results for each of the specified items.</returns>
        private ServiceResponseCollection<ServiceResponse> InternalCreateItems(
            IEnumerable<Item> items,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode,
            ServiceErrorHandling errorHandling)
        {
            CreateItemRequest request = new CreateItemRequest(this, errorHandling);

            request.ParentFolderId = parentFolderId;
            request.Items = items;
            request.MessageDisposition = messageDisposition;
            request.SendInvitationsMode = sendInvitationsMode;

            return request.Execute();
        }

        /// <summary>
        /// Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact, PostItem, Task and Item.
        /// CreateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to create.</param>
        /// <param name="parentFolderId">The Id of the folder in which to place the newly created items. If null, items are created in their default folders.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsMode">Indicates if and how invitations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <returns>A ServiceResponseCollection providing creation results for each of the specified items.</returns>
        public ServiceResponseCollection<ServiceResponse> CreateItems(
            IEnumerable<Item> items,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode)
        {
            // All items have to be new.
            if (!items.TrueForAll((item) => item.IsNew))
            {
                throw new ServiceValidationException(Strings.CreateItemsDoesNotHandleExistingItems);
            }

            // Make sure that all items do *not* have unprocessed attachments.
            if (!items.TrueForAll((item) => !item.HasUnprocessedAttachmentChanges()))
            {
                throw new ServiceValidationException(Strings.CreateItemsDoesNotAllowAttachments);
            }

            return this.InternalCreateItems(
                items,
                parentFolderId,
                messageDisposition,
                sendInvitationsMode,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Creates an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="item">The item to create.</param>
        /// <param name="parentFolderId">The Id of the folder in which to place the newly created item. If null, the item is created in its default folders.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if item is an EmailMessage instance.</param>
        /// <param name="sendInvitationsMode">Indicates if and how invitations should be sent for item of type Appointment. Required if item is an Appointment instance.</param>
        internal void CreateItem(
            Item item,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode)
        {
            this.InternalCreateItems(
                new Item[] { item },
                parentFolderId,
                messageDisposition,
                sendInvitationsMode,
                ServiceErrorHandling.ThrowOnError);
        }

        /// <summary>
        /// Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <param name="suppressReadReceipt">Whether to suppress read receipts</param>
        /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
        private ServiceResponseCollection<UpdateItemResponse> InternalUpdateItems(
            IEnumerable<Item> items,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            ServiceErrorHandling errorHandling,
            bool suppressReadReceipt)
        {
            UpdateItemRequest request = new UpdateItemRequest(this, errorHandling);

            request.Items.AddRange(items);
            request.SavedItemsDestinationFolder = savedItemsDestinationFolderId;
            request.MessageDisposition = messageDisposition;
            request.ConflictResolutionMode = conflictResolution;
            request.SendInvitationsOrCancellationsMode = sendInvitationsOrCancellationsMode;
            request.SuppressReadReceipts = suppressReadReceipt;

            return request.Execute();
        }

        /// <summary>
        /// Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
        public ServiceResponseCollection<UpdateItemResponse> UpdateItems(
            IEnumerable<Item> items,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode)
        {
            return this.UpdateItems(items, savedItemsDestinationFolderId, conflictResolution, messageDisposition, sendInvitationsOrCancellationsMode, false);
        }

        /// <summary>
        /// Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
        public ServiceResponseCollection<UpdateItemResponse> UpdateItems(
            IEnumerable<Item> items,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            bool suppressReadReceipts)
        {
            // All items have to exist on the server (!new) and modified (dirty)
            if (!items.TrueForAll((item) => (!item.IsNew && item.IsDirty)))
            {
                throw new ServiceValidationException(Strings.UpdateItemsDoesNotSupportNewOrUnchangedItems);
            }

            // Make sure that all items do *not* have unprocessed attachments.
            if (!items.TrueForAll((item) => !item.HasUnprocessedAttachmentChanges()))
            {
                throw new ServiceValidationException(Strings.UpdateItemsDoesNotAllowAttachments);
            }

            return this.InternalUpdateItems(
                items,
                savedItemsDestinationFolderId,
                conflictResolution,
                messageDisposition,
                sendInvitationsOrCancellationsMode,
                ServiceErrorHandling.ReturnErrors,
                suppressReadReceipts);
        }

        /// <summary>
        /// Updates an item.
        /// </summary>
        /// <param name="item">The item to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the message, meeting invitation or cancellation is saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for an item of type EmailMessage. Required if item is an EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for ian tem of type Appointment. Required if item is an Appointment instance.</param>
        /// <returns>Updated item.</returns>
        internal Item UpdateItem(
            Item item,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode)
        {
            return this.UpdateItem(item, savedItemsDestinationFolderId, conflictResolution, messageDisposition, sendInvitationsOrCancellationsMode, false);
        }

        /// <summary>
        /// Updates an item.
        /// </summary>
        /// <param name="item">The item to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the message, meeting invitation or cancellation is saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for an item of type EmailMessage. Required if item is an EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for ian tem of type Appointment. Required if item is an Appointment instance.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>Updated item.</returns>
        internal Item UpdateItem(
            Item item,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            bool suppressReadReceipts)
        {
            ServiceResponseCollection<UpdateItemResponse> responses = this.InternalUpdateItems(
                new Item[] { item },
                savedItemsDestinationFolderId,
                conflictResolution,
                messageDisposition,
                sendInvitationsOrCancellationsMode,
                ServiceErrorHandling.ThrowOnError,
                suppressReadReceipts);

            return responses[0].ReturnedItem;
        }

        /// <summary>
        /// Sends an item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="savedCopyDestinationFolderId">The saved copy destination folder id.</param>
        internal void SendItem(
            Item item,
            FolderId savedCopyDestinationFolderId)
        {
            SendItemRequest request = new SendItemRequest(this, ServiceErrorHandling.ThrowOnError);

            request.Items = new Item[] { item };
            request.SavedCopyDestinationFolderId = savedCopyDestinationFolderId;

            request.Execute();
        }

        /// <summary>
        /// Copies multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        private ServiceResponseCollection<MoveCopyItemResponse> InternalCopyItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool? returnNewItemIds,
            ServiceErrorHandling errorHandling)
        {
            CopyItemRequest request = new CopyItemRequest(this, errorHandling);
            request.ItemIds.AddRange(itemIds);
            request.DestinationFolderId = destinationFolderId;
            request.ReturnNewItemIds = returnNewItemIds;

            return request.Execute();
        }

        /// <summary>
        /// Copies multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<MoveCopyItemResponse> CopyItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId)
        {
            return this.InternalCopyItems(
                itemIds,
                destinationFolderId,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Copies multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<MoveCopyItemResponse> CopyItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool returnNewItemIds)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "CopyItems");

            return this.InternalCopyItems(
                itemIds,
                destinationFolderId,
                returnNewItemIds,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Copies an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="itemId">The Id of the item to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the item to.</param>
        /// <returns>The copy of the item.</returns>
        internal Item CopyItem(
            ItemId itemId,
            FolderId destinationFolderId)
        {
            return this.InternalCopyItems(
                new ItemId[] { itemId },
                destinationFolderId,
                null,
                ServiceErrorHandling.ThrowOnError)[0].Item;
        }

        /// <summary>
        /// Moves multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        private ServiceResponseCollection<MoveCopyItemResponse> InternalMoveItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool? returnNewItemIds,
            ServiceErrorHandling errorHandling)
        {
            MoveItemRequest request = new MoveItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.DestinationFolderId = destinationFolderId;
            request.ReturnNewItemIds = returnNewItemIds;

            return request.Execute();
        }

        /// <summary>
        /// Moves multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<MoveCopyItemResponse> MoveItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId)
        {
            return this.InternalMoveItems(
                itemIds,
                destinationFolderId,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Moves multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<MoveCopyItemResponse> MoveItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool returnNewItemIds)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "MoveItems");

            return this.InternalMoveItems(
                itemIds,
                destinationFolderId,
                returnNewItemIds,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Move an item.
        /// </summary>
        /// <param name="itemId">The Id of the item to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the item to.</param>
        /// <returns>The moved item.</returns>
        internal Item MoveItem(
            ItemId itemId,
            FolderId destinationFolderId)
        {
            return this.InternalMoveItems(
                new ItemId[] { itemId },
                destinationFolderId,
                null,
                ServiceErrorHandling.ThrowOnError)[0].Item;
        }

        /// <summary>
        /// Archives multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="sourceFolderId">The Id of the folder in primary corresponding to which items are being archived to.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<ArchiveItemResponse> ArchiveItems(
            IEnumerable<ItemId> itemIds,
            FolderId sourceFolderId)
        {
            ArchiveItemRequest request = new ArchiveItemRequest(this, ServiceErrorHandling.ReturnErrors);

            request.Ids.AddRange(itemIds);
            request.SourceFolderId = sourceFolderId;

            return request.Execute();
        }

        /// <summary>
        /// Finds items.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="parentFolderIds">The parent folder ids.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="queryString">query string to be used for indexed search.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by.</param>
        /// <param name="errorHandlingMode">Indicates the type of error handling should be done.</param>
        /// <returns>Service response collection.</returns>
        internal ServiceResponseCollection<FindItemResponse<TItem>> FindItems<TItem>(
            IEnumerable<FolderId> parentFolderIds,
            SearchFilter searchFilter,
            string queryString,
            ViewBase view,
            Grouping groupBy,
            ServiceErrorHandling errorHandlingMode)
            where TItem : Item
        {
            EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            FindItemRequest<TItem> request = new FindItemRequest<TItem>(this, errorHandlingMode);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.SearchFilter = searchFilter;
            request.QueryString = queryString;
            request.View = view;
            request.GroupBy = groupBy;

            return request.Execute();
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">the search string to be used for indexed search, if any.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(FolderId parentFolderId, string queryString, ViewBase view)
        {
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");

            ServiceResponseCollection<FindItemResponse<Item>> responses = this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                queryString,
                view,
                null,   /* groupBy */
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. 
        /// Along with conversations, a list of highlight terms are returned.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">the search string to be used for indexed search, if any.</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(FolderId parentFolderId, string queryString, bool returnHighlightTerms, ViewBase view)
        {
            FolderId[] parentFolderIds = new FolderId[] { parentFolderId };

            EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParamAllowNull(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "FindItems");

            FindItemRequest<Item> request = new FindItemRequest<Item>(this, ServiceErrorHandling.ThrowOnError);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;
            request.View = view;

            ServiceResponseCollection<FindItemResponse<Item>> responses = request.Execute();
            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. 
        /// Along with conversations, a list of highlight terms are returned.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">the search string to be used for indexed search, if any.</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public GroupedFindItemsResults<Item> FindItems(FolderId parentFolderId, string queryString, bool returnHighlightTerms, ViewBase view, Grouping groupBy)
        {
            FolderId[] parentFolderIds = new FolderId[] { parentFolderId };

            EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParamAllowNull(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "FindItems");

            FindItemRequest<Item> request = new FindItemRequest<Item>(this, ServiceErrorHandling.ThrowOnError);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;
            request.View = view;
            request.GroupBy = groupBy;

            ServiceResponseCollection<FindItemResponse<Item>> responses = request.Execute();
            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(FolderId parentFolderId, SearchFilter searchFilter, ViewBase view)
        {
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindItemResponse<Item>> responses = this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                searchFilter,
                null, /* queryString */
                view,
                null,   /* groupBy */
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(FolderId parentFolderId, ViewBase view)
        {
            ServiceResponseCollection<FindItemResponse<Item>> responses = this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                null, /* queryString */
                view,
                null, /* groupBy */
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(WellKnownFolderName parentFolderName, string queryString, ViewBase view)
        {
            return this.FindItems(new FolderId(parentFolderName), queryString, view);
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(WellKnownFolderName parentFolderName, SearchFilter searchFilter, ViewBase view)
        {
            return this.FindItems(
                new FolderId(parentFolderName),
                searchFilter,
                view);
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public FindItemsResults<Item> FindItems(WellKnownFolderName parentFolderName, ViewBase view)
        {
            return this.FindItems(
                new FolderId(parentFolderName),
                (SearchFilter)null,
                view);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        public GroupedFindItemsResults<Item> FindItems(
            FolderId parentFolderId,
            string queryString,
            ViewBase view,
            Grouping groupBy)
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");

            ServiceResponseCollection<FindItemResponse<Item>> responses = this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                queryString,
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        public GroupedFindItemsResults<Item> FindItems(
            FolderId parentFolderId,
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy)
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindItemResponse<Item>> responses = this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                searchFilter,
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        public GroupedFindItemsResults<Item> FindItems(
            FolderId parentFolderId,
            ViewBase view,
            Grouping groupBy)
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");

            ServiceResponseCollection<FindItemResponse<Item>> responses = this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <typeparam name="TItem">Type of item.</typeparam>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        internal ServiceResponseCollection<FindItemResponse<TItem>> FindItems<TItem>(
            FolderId parentFolderId,
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy)
            where TItem : Item
        {
            return this.FindItems<TItem>(
                new FolderId[] { parentFolderId },
                searchFilter,
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A collection of grouped items representing the contents of the specified.</returns>
        public GroupedFindItemsResults<Item> FindItems(
            WellKnownFolderName parentFolderName,
            string queryString,
            ViewBase view,
            Grouping groupBy)
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");

            return this.FindItems(
                new FolderId(parentFolderName),
                queryString,
                view,
                groupBy);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A collection of grouped items representing the contents of the specified.</returns>
        public GroupedFindItemsResults<Item> FindItems(
            WellKnownFolderName parentFolderName,
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy)
        {
            return this.FindItems(
                new FolderId(parentFolderName),
                searchFilter,
                view,
                groupBy);
        }

        /// <summary>
        /// Obtains a list of appointments by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The id of the calendar folder in which to search for items.</param>
        /// <param name="calendarView">The calendar view controlling the number of appointments returned.</param>
        /// <returns>A collection of appointments representing the contents of the specified folder.</returns>
        public FindItemsResults<Appointment> FindAppointments(FolderId parentFolderId, CalendarView calendarView)
        {
            ServiceResponseCollection<FindItemResponse<Appointment>> response = this.FindItems<Appointment>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                null, /* queryString */
                calendarView,
                null, /* groupBy */
                ServiceErrorHandling.ThrowOnError);

            return response[0].Results;
        }

        /// <summary>
        /// Obtains a list of appointments by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the calendar folder in which to search for items.</param>
        /// <param name="calendarView">The calendar view controlling the number of appointments returned.</param>
        /// <returns>A collection of appointments representing the contents of the specified folder.</returns>
        public FindItemsResults<Appointment> FindAppointments(WellKnownFolderName parentFolderName, CalendarView calendarView)
        {
            return this.FindAppointments(new FolderId(parentFolderName), calendarView);
        }

        /// <summary>
        /// Loads the properties of multiple items in a single call to EWS.
        /// </summary>
        /// <param name="items">The items to load the properties of.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified items.</returns>
        public ServiceResponseCollection<ServiceResponse> LoadPropertiesForItems(IEnumerable<Item> items, PropertySet propertySet)
        {
            EwsUtilities.ValidateParamCollection(items, "items");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            return this.InternalLoadPropertiesForItems(
                items,
                propertySet,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Loads the properties of multiple items in a single call to EWS.
        /// </summary>
        /// <param name="items">The items to load the properties of.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="errorHandling">Indicates the type of error handling should be done.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified items.</returns>
        internal ServiceResponseCollection<ServiceResponse> InternalLoadPropertiesForItems(
            IEnumerable<Item> items,
            PropertySet propertySet,
            ServiceErrorHandling errorHandling)
        {
            GetItemRequestForLoad request = new GetItemRequestForLoad(this, errorHandling);

            request.ItemIds.AddRange(items);
            request.PropertySet = propertySet;

            return request.Execute();
        }

        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        private ServiceResponseCollection<GetItemResponse> InternalBindToItems(
            IEnumerable<ItemId> itemIds,
            PropertySet propertySet,
            string anchorMailbox,
            ServiceErrorHandling errorHandling)
        {
            GetItemRequest request = new GetItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.PropertySet = propertySet;
            request.AnchorMailbox = anchorMailbox;

            return request.Execute();
        }

        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<GetItemResponse> BindToItems(IEnumerable<ItemId> itemIds, PropertySet propertySet)
        {
            EwsUtilities.ValidateParamCollection(itemIds, "itemIds");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            return this.InternalBindToItems(
                itemIds,
                propertySet,
                null, /* anchorMailbox */
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        /// <remarks>
        /// This API designed to be used primarily in groups scenarios where we want to set the
        /// anchor mailbox header so that request is routed directly to the group mailbox backend server.
        /// </remarks>
        public ServiceResponseCollection<GetItemResponse> BindToGroupItems(
            IEnumerable<ItemId> itemIds,
            PropertySet propertySet,
            string anchorMailbox)
        {
            EwsUtilities.ValidateParamCollection(itemIds, "itemIds");
            EwsUtilities.ValidateParam(propertySet, "propertySet");
            EwsUtilities.ValidateParam(propertySet, "anchorMailbox");

            return this.InternalBindToItems(
                itemIds,
                propertySet,
                anchorMailbox,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Binds to item.
        /// </summary>
        /// <param name="itemId">The item id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Item.</returns>
        internal Item BindToItem(ItemId itemId, PropertySet propertySet)
        {
            EwsUtilities.ValidateParam(itemId, "itemId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            ServiceResponseCollection<GetItemResponse> responses = this.InternalBindToItems(
                new ItemId[] { itemId },
                propertySet,
                null, /* anchorMailbox */
                ServiceErrorHandling.ThrowOnError);

            return responses[0].Item;
        }

        /// <summary>
        /// Binds to item.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="itemId">The item id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Item</returns>
        internal TItem BindToItem<TItem>(ItemId itemId, PropertySet propertySet)
            where TItem : Item
        {
            Item result = this.BindToItem(itemId, propertySet);

            if (result is TItem)
            {
                return (TItem)result;
            }
            else
            {
                throw new ServiceLocalException(
                    string.Format(
                        Strings.ItemTypeNotCompatible,
                        result.GetType().Name,
                        typeof(TItem).Name));
            }
        }

        /// <summary>
        /// Deletes multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if any of the item Ids represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if any of the item Ids represents a Task.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
        private ServiceResponseCollection<ServiceResponse> InternalDeleteItems(
            IEnumerable<ItemId> itemIds,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            ServiceErrorHandling errorHandling,
            bool suppressReadReceipts)
        {
            DeleteItemRequest request = new DeleteItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.DeleteMode = deleteMode;
            request.SendCancellationsMode = sendCancellationsMode;
            request.AffectedTaskOccurrences = affectedTaskOccurrences;
            request.SuppressReadReceipts = suppressReadReceipts;

            return request.Execute();
        }

        /// <summary>
        /// Deletes multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if any of the item Ids represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if any of the item Ids represents a Task.</param>
        /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
        public ServiceResponseCollection<ServiceResponse> DeleteItems(
            IEnumerable<ItemId> itemIds,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences)
        {
            return this.DeleteItems(itemIds, deleteMode, sendCancellationsMode, affectedTaskOccurrences, false);
        }

        /// <summary>
        /// Deletes multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if any of the item Ids represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if any of the item Ids represents a Task.</param>
        /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
        /// <param name="suppressReadReceipt">Whether to suppress read receipts</param>
        public ServiceResponseCollection<ServiceResponse> DeleteItems(
            IEnumerable<ItemId> itemIds,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            bool suppressReadReceipt)
        {
            EwsUtilities.ValidateParamCollection(itemIds, "itemIds");

            return this.InternalDeleteItems(
                itemIds,
                deleteMode,
                sendCancellationsMode,
                affectedTaskOccurrences,
                ServiceErrorHandling.ReturnErrors,
                suppressReadReceipt);
        }

        /// <summary>
        /// Deletes an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="itemId">The Id of the item to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if the item Id represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if item Id represents a Task.</param>
        internal void DeleteItem(
            ItemId itemId,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences)
        {
            this.DeleteItem(itemId, deleteMode, sendCancellationsMode, affectedTaskOccurrences, false);
        }

        /// <summary>
        /// Deletes an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="itemId">The Id of the item to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if the item Id represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if item Id represents a Task.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        internal void DeleteItem(
            ItemId itemId,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            bool suppressReadReceipts)
        {
            EwsUtilities.ValidateParam(itemId, "itemId");

            this.InternalDeleteItems(
                new ItemId[] { itemId },
                deleteMode,
                sendCancellationsMode,
                affectedTaskOccurrences,
                ServiceErrorHandling.ThrowOnError,
                suppressReadReceipts);
        }

        /// <summary>
        /// Mark items as junk.
        /// </summary>
        /// <param name="itemIds">ItemIds for the items to mark</param>
        /// <param name="isJunk">Whether the items are junk.  If true, senders are add to blocked sender list. If false, senders are removed.</param>
        /// <param name="moveItem">Whether to move the item.  Items are moved to junk folder if isJunk is true, inbox if isJunk is false.</param>
        /// <returns>A ServiceResponseCollection providing itemIds for each of the moved items..</returns>
        public ServiceResponseCollection<MarkAsJunkResponse> MarkAsJunk(IEnumerable<ItemId> itemIds, bool isJunk, bool moveItem)
        {
            MarkAsJunkRequest request = new MarkAsJunkRequest(this, ServiceErrorHandling.ReturnErrors);
            request.ItemIds.AddRange(itemIds);
            request.IsJunk = isJunk;
            request.MoveItem = moveItem;
            return request.Execute();
        }

        #endregion

        #region People operations

        /// <summary>
        /// This method is for search scenarios. Retrieves a set of personas satisfying the specified search conditions.
        /// </summary>
        /// <param name="folderId">Id of the folder being searched</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view which defines the number of persona being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>A collection of personas matching the search conditions</returns>
        public ICollection<Persona> FindPeople(FolderId folderId, SearchFilter searchFilter, ViewBase view, string queryString)
        {
            EwsUtilities.ValidateParamAllowNull(folderId, "folderId");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(queryString, "queryString");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013_SP1, "FindPeople");

            FindPeopleRequest request = new FindPeopleRequest(this);

            request.FolderId = folderId;
            request.SearchFilter = searchFilter;
            request.View = view;
            request.QueryString = queryString;

            return request.Execute().Personas;
        }

        /// <summary>
        /// This method is for search scenarios. Retrieves a set of personas satisfying the specified search conditions.
        /// </summary>
        /// <param name="folderName">Name of the folder being searched</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view which defines the number of persona being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>A collection of personas matching the search conditions</returns>
        public ICollection<Persona> FindPeople(WellKnownFolderName folderName, SearchFilter searchFilter, ViewBase view, string queryString)
        {
            return this.FindPeople(new FolderId(folderName), searchFilter, view, queryString);
        }

        /// <summary>
        /// This method is for browse scenarios. Retrieves a set of personas satisfying the specified browse conditions.
        /// Browse scenariosdon't require query string.
        /// </summary>
        /// <param name="folderId">Id of the folder being browsed</param>
        /// <param name="searchFilter">Search filter</param>
        /// <param name="view">The view which defines paging and the number of persona being returned</param>
        /// <returns>A result object containing resultset for browsing</returns>
        public FindPeopleResults FindPeople(FolderId folderId, SearchFilter searchFilter, ViewBase view)
        {
            EwsUtilities.ValidateParamAllowNull(folderId, "folderId");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");
            EwsUtilities.ValidateParamAllowNull(view, "view");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013_SP1, "FindPeople");

            FindPeopleRequest request = new FindPeopleRequest(this);

            request.FolderId = folderId;
            request.SearchFilter = searchFilter;
            request.View = view;

            return request.Execute().Results;
        }

        /// <summary>
        /// This method is for browse scenarios. Retrieves a set of personas satisfying the specified browse conditions.
        /// Browse scenarios don't require query string.
        /// </summary>
        /// <param name="folderName">Name of the folder being browsed</param>
        /// <param name="searchFilter">Search filter</param>
        /// <param name="view">The view which defines paging and the number of personas being returned</param>
        /// <returns>A result object containing resultset for browsing</returns>
        public FindPeopleResults FindPeople(WellKnownFolderName folderName, SearchFilter searchFilter, ViewBase view)
        {
            return this.FindPeople(new FolderId(folderName), searchFilter, view);
        }

        /// <summary>
        /// Retrieves all people who are relevant to the user
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <returns>A collection of personas matching the query string</returns>
        public IPeopleQueryResults BrowsePeople(ViewBase view)
        {
            return this.BrowsePeople(view, null);
        }

        /// <summary>
        /// Retrieves all people who are relevant to the user
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="context">The context for this query. See PeopleQueryContextKeys for keys</param>
        /// <returns>A collection of personas matching the query string</returns>
        public IPeopleQueryResults BrowsePeople(ViewBase view, Dictionary<string, string> context)
        {
            return this.PerformPeopleQuery(view, string.Empty, context, null);
        }

        /// <summary>
        /// Searches for people who are relevant to the user, automatically determining
        /// the best sources to use.
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>A collection of personas matching the query string</returns>
        public IPeopleQueryResults SearchPeople(ViewBase view, string queryString)
        {
            return this.SearchPeople(view, queryString, null, null);
        }

        /// <summary>
        /// Searches for people who are relevant to the user
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="context">The context for this query. See PeopleQueryContextKeys for keys</param>
        /// <param name="queryMode">The scope of the query.</param>
        /// <returns>A collection of personas matching the query string</returns>
        public IPeopleQueryResults SearchPeople(ViewBase view, string queryString, Dictionary<string, string> context, PeopleQueryMode queryMode)
        {
            EwsUtilities.ValidateParam(queryString, "queryString");

            return this.PerformPeopleQuery(view, queryString, context, queryMode);
        }

        /// <summary>
        /// Performs a People Query FindPeople call
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="context">The context for this query</param>
        /// <param name="queryMode">The scope of the query.</param>
        /// <returns></returns>
        private IPeopleQueryResults PerformPeopleQuery(ViewBase view, string queryString, Dictionary<string, string> context, PeopleQueryMode queryMode)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2015, "FindPeople");

            if (context == null)
            {
                context = new Dictionary<string, string>();
            }

            if (queryMode == null)
            {
                queryMode = PeopleQueryMode.Auto;
            }

            FindPeopleRequest request = new FindPeopleRequest(this);
            request.View = view;
            request.QueryString = queryString;
            request.SearchPeopleSuggestionIndex = true;
            request.Context = context;
            request.QueryMode = queryMode;

            FindPeopleResponse response = request.Execute();

            PeopleQueryResults results = new PeopleQueryResults();
            results.Personas = response.Personas.ToList();
            results.TransactionId = response.TransactionId;

            return results;
        }

        /// <summary>
        /// Get a user's photo.
        /// </summary>
        /// <param name="emailAddress">The user's email address</param>
        /// <param name="userPhotoSize">The desired size of the returned photo. Valid photo sizes are in UserPhotoSize</param>
        /// <param name="entityTag">A photo's cache ID which will allow the caller to ensure their cached photo is up to date</param>
        /// <returns>A result object containing the photo state</returns>
        public GetUserPhotoResults GetUserPhoto(string emailAddress, string userPhotoSize, string entityTag)
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");
            EwsUtilities.ValidateParam(userPhotoSize, "userPhotoSize");
            EwsUtilities.ValidateParamAllowNull(entityTag, "entityTag");

            GetUserPhotoRequest request = new GetUserPhotoRequest(this);

            request.EmailAddress = emailAddress;
            request.UserPhotoSize = userPhotoSize;
            request.EntityTag = entityTag;

            return request.Execute().Results;
        }

        /// <summary>
        /// Set a user's photo.
        /// </summary>
        /// <param name="emailAddress">The user's email address</param>
        /// <param name="photo">The photo to set</param>
        /// <returns>A result object</returns>
        public SetUserPhotoResults SetUserPhoto(string emailAddress, byte[] photo)
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");
            EwsUtilities.ValidateParam(photo, "photo");
            
            SetUserPhotoRequest request = new SetUserPhotoRequest(this);

            request.EmailAddress = emailAddress;
            request.Photo = photo;

            return request.Execute().Results;
        }

        /// <summary>
        /// Begins an async request for a user photo
        /// </summary>
        /// <param name="callback">An AsyncCallback delegate</param>
        /// <param name="state">An object that contains state information for this request</param>
        /// <param name="emailAddress">The user's email address</param>
        /// <param name="userPhotoSize">The desired size of the returned photo. Valid photo sizes are in UserPhotoSize</param>
        /// <param name="entityTag">A photo's cache ID which will allow the caller to ensure their cached photo is up to date</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginGetUserPhoto(
            AsyncCallback callback,
            object state,
            string emailAddress,
            string userPhotoSize,
            string entityTag)
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");
            EwsUtilities.ValidateParam(userPhotoSize, "userPhotoSize");
            EwsUtilities.ValidateParamAllowNull(entityTag, "entityTag");

            GetUserPhotoRequest request = new GetUserPhotoRequest(this);

            request.EmailAddress = emailAddress;
            request.UserPhotoSize = userPhotoSize;
            request.EntityTag = entityTag;

            return request.BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an async request for a user's photo
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A result object containing the photo state</returns>
        public GetUserPhotoResults EndGetUserPhoto(IAsyncResult asyncResult)
        {
            GetUserPhotoRequest request = AsyncRequestResult.ExtractServiceRequest<GetUserPhotoRequest>(this, asyncResult);
            return request.EndExecute(asyncResult).Results;
        }

        #endregion

        #region PeopleInsights operations

        /// <summary>
        /// This method is for retreiving people insight for given email addresses
        /// </summary>
        /// <param name="emailAddresses">Specified eamiladdresses to retrieve</param>
        /// <returns>The collection of Person objects containing the insight info</returns>
        public Collection<Person> GetPeopleInsights(IEnumerable<string> emailAddresses)
        {
            GetPeopleInsightsRequest request = new GetPeopleInsightsRequest(this);
            request.Emailaddresses.AddRange(emailAddresses);

            return request.Execute().People;
        }

        #endregion
        #region Attachment operations

        /// <summary>
        /// Gets an attachment.
        /// </summary>
        /// <param name="attachments">The attachments.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>Service response collection.</returns>
        private ServiceResponseCollection<GetAttachmentResponse> InternalGetAttachments(
            IEnumerable<Attachment> attachments,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties,
            ServiceErrorHandling errorHandling)
        {
            GetAttachmentRequest request = new GetAttachmentRequest(this, errorHandling);

            request.Attachments.AddRange(attachments);
            request.BodyType = bodyType;

            if (additionalProperties != null)
            {
                request.AdditionalProperties.AddRange(additionalProperties);
            }

            return request.Execute();
        }

        /// <summary>
        /// Gets attachments.
        /// </summary>
        /// <param name="attachments">The attachments.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        /// <returns>Service response collection.</returns>
        public ServiceResponseCollection<GetAttachmentResponse> GetAttachments(
            Attachment[] attachments,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            return this.InternalGetAttachments(
                attachments,
                bodyType,
                additionalProperties,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Gets attachments.
        /// </summary>
        /// <param name="attachmentIds">The attachment ids.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        /// <returns>Service response collection.</returns>
        public ServiceResponseCollection<GetAttachmentResponse> GetAttachments(
            string[] attachmentIds,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            GetAttachmentRequest request = new GetAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

            request.AttachmentIds.AddRange(attachmentIds);
            request.BodyType = bodyType;

            if (additionalProperties != null)
            {
                request.AdditionalProperties.AddRange(additionalProperties);
            }

            return request.Execute();
        }

        /// <summary>
        /// Gets an attachment.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        internal void GetAttachment(
            Attachment attachment,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            this.InternalGetAttachments(
                new Attachment[] { attachment },
                bodyType,
                additionalProperties,
                ServiceErrorHandling.ThrowOnError);
        }

        /// <summary>
        /// Creates attachments.
        /// </summary>
        /// <param name="parentItemId">The parent item id.</param>
        /// <param name="attachments">The attachments.</param>
        /// <returns>Service response collection.</returns>
        internal ServiceResponseCollection<CreateAttachmentResponse> CreateAttachments(
            string parentItemId,
            IEnumerable<Attachment> attachments)
        {
            CreateAttachmentRequest request = new CreateAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

            request.ParentItemId = parentItemId;
            request.Attachments.AddRange(attachments);

            return request.Execute();
        }

        /// <summary>
        /// Deletes attachments.
        /// </summary>
        /// <param name="attachments">The attachments.</param>
        /// <returns>Service response collection.</returns>
        internal ServiceResponseCollection<DeleteAttachmentResponse> DeleteAttachments(IEnumerable<Attachment> attachments)
        {
            DeleteAttachmentRequest request = new DeleteAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

            request.Attachments.AddRange(attachments);

            return request.Execute();
        }

        #endregion

        #region AD related operations

        /// <summary>
        /// Finds contacts in the user's Contacts folder and the Global Address List (in that order) that have names
        /// that match the one passed as a parameter. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public NameResolutionCollection ResolveName(string nameToResolve)
        {
            return this.ResolveName(
                nameToResolve,
                ResolveNameSearchLocation.ContactsThenDirectory,
                false);
        }

        /// <summary>
        /// Finds contacts in the Global Address List and/or in specific contact folders that have names
        /// that match the one passed as a parameter. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="parentFolderIds">The Ids of the contact folders in which to look for matching contacts.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public NameResolutionCollection ResolveName(
            string nameToResolve,
            IEnumerable<FolderId> parentFolderIds,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails)
        {
            return ResolveName(
                nameToResolve,
                parentFolderIds,
                searchScope,
                returnContactDetails,
                null);
        }

        /// <summary>
        /// Finds contacts in the Global Address List and/or in specific contact folders that have names
        /// that match the one passed as a parameter. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="parentFolderIds">The Ids of the contact folders in which to look for matching contacts.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <param name="contactDataPropertySet">The property set for the contct details</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public NameResolutionCollection ResolveName(
            string nameToResolve,
            IEnumerable<FolderId> parentFolderIds,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails,
            PropertySet contactDataPropertySet)
        {
            if (contactDataPropertySet != null)
            {
                EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "ResolveName");
            }

            EwsUtilities.ValidateParam(nameToResolve, "nameToResolve");
            if (parentFolderIds != null)
            {
                EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            }

            ResolveNamesRequest request = new ResolveNamesRequest(this);

            request.NameToResolve = nameToResolve;
            request.ReturnFullContactData = returnContactDetails;
            request.ParentFolderIds.AddRange(parentFolderIds);
            request.SearchLocation = searchScope;
            request.ContactDataPropertySet = contactDataPropertySet;

            return request.Execute()[0].Resolutions;
        }

        /// <summary>
        /// Finds contacts in the Global Address List that have names that match the one passed as a parameter.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <param name="contactDataPropertySet">Propety set for contact details</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public NameResolutionCollection ResolveName(
            string nameToResolve,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails,
            PropertySet contactDataPropertySet)
        {
            return this.ResolveName(
                nameToResolve,
                null,
                searchScope,
                returnContactDetails,
                contactDataPropertySet);
        }

        /// <summary>
        /// Finds contacts in the Global Address List that have names that match the one passed as a parameter.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public NameResolutionCollection ResolveName(
            string nameToResolve,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails)
        {
            return this.ResolveName(
                nameToResolve,
                null,
                searchScope,
                returnContactDetails);
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="emailAddress">The e-mail address of the group.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public ExpandGroupResults ExpandGroup(EmailAddress emailAddress)
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");

            ExpandGroupRequest request = new ExpandGroupRequest(this);

            request.EmailAddress = emailAddress;

            return request.Execute()[0].Members;
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="groupId">The Id of the group to expand.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public ExpandGroupResults ExpandGroup(ItemId groupId)
        {
            EwsUtilities.ValidateParam(groupId, "groupId");

            EmailAddress emailAddress = new EmailAddress();
            emailAddress.Id = groupId;

            return this.ExpandGroup(emailAddress);
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the group to expand.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public ExpandGroupResults ExpandGroup(string smtpAddress)
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");

            return this.ExpandGroup(new EmailAddress(smtpAddress));
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="address">The SMTP address of the group to expand.</param>
        /// <param name="routingType">The routing type of the address of the group to expand.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public ExpandGroupResults ExpandGroup(string address, string routingType)
        {
            EwsUtilities.ValidateParam(address, "address");
            EwsUtilities.ValidateParam(routingType, "routingType");

            EmailAddress emailAddress = new EmailAddress(address);
            emailAddress.RoutingType = routingType;

            return this.ExpandGroup(emailAddress);
        }

        /// <summary>
        /// Get the password expiration date
        /// </summary>
        /// <param name="mailboxSmtpAddress">The e-mail address of the user.</param>
        /// <returns>The password expiration date.</returns>
        public DateTime? GetPasswordExpirationDate(string mailboxSmtpAddress)
        {
            GetPasswordExpirationDateRequest request = new GetPasswordExpirationDateRequest(this);
            request.MailboxSmtpAddress = mailboxSmtpAddress;

            return request.Execute().PasswordExpirationDate;
        }
        #endregion

        #region Notification operations

        /// <summary>
        /// Subscribes to pull notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PullSubscription representing the new subscription.</returns>
        public PullSubscription SubscribeToPullNotifications(
            IEnumerable<FolderId> folderIds,
            int timeout,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToPullNotificationsRequest(
                 folderIds,
                 timeout,
                 watermark,
                 eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribes to pull notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToPullNotifications(
            AsyncCallback callback,
            object state,
            IEnumerable<FolderId> folderIds,
            int timeout,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToPullNotificationsRequest(
                folderIds,
                timeout,
                watermark,
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Subscribes to pull notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PullSubscription representing the new subscription.</returns>
        public PullSubscription SubscribeToPullNotificationsOnAllFolders(
            int timeout,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "SubscribeToPullNotificationsOnAllFolders");

            return this.BuildSubscribeToPullNotificationsRequest(
                null,
                timeout,
                watermark,
                eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to pull notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToPullNotificationsOnAllFolders(
            AsyncCallback callback,
            object state,
            int timeout,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "BeginSubscribeToPullNotificationsOnAllFolders");

            return this.BuildSubscribeToPullNotificationsRequest(
                null,
                timeout,
                watermark,
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an asynchronous request to subscribe to pull notifications in the authenticated user's mailbox. 
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A PullSubscription representing the new subscription.</returns>
        public PullSubscription EndSubscribeToPullNotifications(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<SubscribeToPullNotificationsRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Subscription;
        }

        /// <summary>
        /// Builds a request to subscribe to pull notifications in the authenticated user's mailbox. 
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A request to subscribe to pull notifications in the authenticated user's mailbox. </returns>
        private SubscribeToPullNotificationsRequest BuildSubscribeToPullNotificationsRequest(
            IEnumerable<FolderId> folderIds,
            int timeout,
            string watermark,
            EventType[] eventTypes)
        {
            if (timeout < 1 || timeout > 1440)
            {
                throw new ArgumentOutOfRangeException("timeout", Strings.TimeoutMustBeBetween1And1440);
            }

            EwsUtilities.ValidateParamCollection(eventTypes, "eventTypes");

            SubscribeToPullNotificationsRequest request = new SubscribeToPullNotificationsRequest(this);

            if (folderIds != null)
            {
                request.FolderIds.AddRange(folderIds);
            }

            request.Timeout = timeout;
            request.EventTypes.AddRange(eventTypes);
            request.Watermark = watermark;

            return request;
        }

        /// <summary>
        /// Unsubscribes from a subscription. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="subscriptionId">The Id of the pull subscription to unsubscribe from.</param>
        internal void Unsubscribe(string subscriptionId)
        {
            this.BuildUnsubscribeRequest(subscriptionId).Execute();
        }

        /// <summary>
        /// Begins an asynchronous request to unsubscribe from a subscription. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="subscriptionId">The Id of the pull subscription to unsubscribe from.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        internal IAsyncResult BeginUnsubscribe(
            AsyncCallback callback,
            object state,
            string subscriptionId)
        {
            return this.BuildUnsubscribeRequest(subscriptionId).BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an asynchronous request to unsubscribe from a subscription.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        internal void EndUnsubscribe(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<UnsubscribeRequest>(this, asyncResult);

            request.EndExecute(asyncResult);
        }

        /// <summary>
        /// Buids a request to unsubscribe from a subscription.
        /// </summary>
        /// <param name="subscriptionId">The Id of the subscription for which to get the events.</param>
        /// <returns>A request to unsubscribe from a subscription.</returns>
        private UnsubscribeRequest BuildUnsubscribeRequest(string subscriptionId)
        {
            EwsUtilities.ValidateParam(subscriptionId, "subscriptionId");

            UnsubscribeRequest request = new UnsubscribeRequest(this);

            request.SubscriptionId = subscriptionId;

            return request;
        }

        /// <summary>
        /// Retrieves the latests events associated with a pull subscription. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
        /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
        /// <returns>A GetEventsResults containing a list of events associated with the subscription.</returns>
        internal GetEventsResults GetEvents(string subscriptionId, string watermark)
        {
            return this.BuildGetEventsRequest(subscriptionId, watermark).Execute()[0].Results;
        }

        /// <summary>
        /// Begins an asynchronous request to retrieve the latests events associated with a pull subscription. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
        /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        internal IAsyncResult BeginGetEvents(
            AsyncCallback callback,
            object state,
            string subscriptionId,
            string watermark)
        {
            return this.BuildGetEventsRequest(subscriptionId, watermark).BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an asynchronous request to retrieve the latests events associated with a pull subscription.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A GetEventsResults containing a list of events associated with the subscription.</returns>
        internal GetEventsResults EndGetEvents(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<GetEventsRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Results;
        }

        /// <summary>
        /// Builds an request to retrieve the latests events associated with a pull subscription.
        /// </summary>
        /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
        /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
        /// <returns>An request to retrieve the latests events associated with a pull subscription. </returns>
        private GetEventsRequest BuildGetEventsRequest(
            string subscriptionId,
            string watermark)
        {
            EwsUtilities.ValidateParam(subscriptionId, "subscriptionId");
            EwsUtilities.ValidateParam(watermark, "watermark");

            GetEventsRequest request = new GetEventsRequest(this);

            request.SubscriptionId = subscriptionId;
            request.Watermark = watermark;

            return request;
        }

        /// <summary>
        /// Subscribes to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription SubscribeToPushNotifications(
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                null,
                null, // AnchorMailbox
                eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToPushNotifications(
            AsyncCallback callback,
            object state,
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                null,
                null, // AnchorMailbox
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Subscribes to push notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription SubscribeToPushNotificationsOnAllFolders(
            Uri url,
            int frequency,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "SubscribeToPushNotificationsOnAllFolders");

            return this.BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                null,
                null, // AnchorMailbox
                eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to push notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="url"></param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToPushNotificationsOnAllFolders(
            AsyncCallback callback,
            object state,
            Uri url,
            int frequency,
            string watermark,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "BeginSubscribeToPushNotificationsOnAllFolders");

            return this.BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                null,
                null, // AnchorMailbox
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Subscribes to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription SubscribeToPushNotifications(
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToPushNotifications(
            AsyncCallback callback,
            object state,
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Subscribes to push notifications on a group mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="groupMailboxSmtp">The smtpaddress of the group mailbox to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription SubscribeToGroupPushNotifications(
            string groupMailboxSmtp,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            params EventType[] eventTypes)
        {
            var folderIds = new FolderId[] { new FolderId(WellKnownFolderName.Inbox, new Mailbox(groupMailboxSmtp)) };
            return this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                groupMailboxSmtp, // AnchorMailbox
                eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="groupMailboxSmtp">The smtpaddress of the group mailbox to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToGroupPushNotifications(
            AsyncCallback callback,
            object state,
            string groupMailboxSmtp,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            params EventType[] eventTypes)
        {
            var folderIds = new FolderId[] { new FolderId(WellKnownFolderName.Inbox, new Mailbox(groupMailboxSmtp)) };
            return this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                groupMailboxSmtp, // AnchorMailbox
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Subscribes to push notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription SubscribeToPushNotificationsOnAllFolders(
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "SubscribeToPushNotificationsOnAllFolders");

            return this.BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to push notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="url"></param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToPushNotificationsOnAllFolders(
            AsyncCallback callback,
            object state,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "BeginSubscribeToPushNotificationsOnAllFolders");

            return this.BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an asynchronous request to subscribe to push notifications in the authenticated user's mailbox.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription EndSubscribeToPushNotifications(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<SubscribeToPushNotificationsRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Subscription;
        }

        /// <summary>
        /// Ends an asynchronous request to subscribe to push notifications in a group mailbox.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public PushSubscription EndSubscribeToGroupPushNotifications(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<SubscribeToPushNotificationsRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Subscription;
        }

        /// <summary>
        /// Set a TeamMailbox
        /// </summary>
        /// <param name="emailAddress">TeamMailbox email address</param>
        /// <param name="sharePointSiteUrl">SharePoint site URL</param>
        /// <param name="state">TeamMailbox lifecycle state</param>
        public void SetTeamMailbox(EmailAddress emailAddress, Uri sharePointSiteUrl, TeamMailboxLifecycleState state)
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetTeamMailbox");

            if (emailAddress == null)
            {
                throw new ArgumentNullException("emailAddress");
            }

            if (sharePointSiteUrl == null)
            {
                throw new ArgumentNullException("sharePointSiteUrl");
            }

            SetTeamMailboxRequest request = new SetTeamMailboxRequest(this, emailAddress, sharePointSiteUrl, state);
            request.Execute();
        }

        /// <summary>
        /// Unpin a TeamMailbox
        /// </summary>
        /// <param name="emailAddress">TeamMailbox email address</param>
        public void UnpinTeamMailbox(EmailAddress emailAddress)
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "UnpinTeamMailbox");

            if (emailAddress == null)
            {
                throw new ArgumentNullException("emailAddress");
            }

            UnpinTeamMailboxRequest request = new UnpinTeamMailboxRequest(this, emailAddress);
            request.Execute();
        }

        /// <summary>
        /// Builds an request to request to subscribe to push notifications in the authenticated user's mailbox.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="anchorMailbox">The smtpaddress of the mailbox to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A request to request to subscribe to push notifications in the authenticated user's mailbox.</returns>
        private SubscribeToPushNotificationsRequest BuildSubscribeToPushNotificationsRequest(
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            string anchorMailbox,
            EventType[] eventTypes)
        {
            EwsUtilities.ValidateParam(url, "url");

            if (frequency < 1 || frequency > 1440)
            {
                throw new ArgumentOutOfRangeException("frequency", Strings.FrequencyMustBeBetween1And1440);
            }

            EwsUtilities.ValidateParamCollection(eventTypes, "eventTypes");

            SubscribeToPushNotificationsRequest request = new SubscribeToPushNotificationsRequest(this);
            request.AnchorMailbox = anchorMailbox;

            if (folderIds != null)
            {
                request.FolderIds.AddRange(folderIds);
            }

            request.Url = url;
            request.Frequency = frequency;
            request.EventTypes.AddRange(eventTypes);
            request.Watermark = watermark;
            request.CallerData = callerData;

            return request;
        }

        /// <summary>
        /// Subscribes to streaming notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A StreamingSubscription representing the new subscription.</returns>
        public StreamingSubscription SubscribeToStreamingNotifications(
            IEnumerable<FolderId> folderIds,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "SubscribeToStreamingNotifications");

            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToStreamingNotificationsRequest(folderIds, eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to streaming notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToStreamingNotifications(
            AsyncCallback callback,
            object state,
            IEnumerable<FolderId> folderIds,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "BeginSubscribeToStreamingNotifications");

            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return this.BuildSubscribeToStreamingNotificationsRequest(folderIds, eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Subscribes to streaming notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A StreamingSubscription representing the new subscription.</returns>
        public StreamingSubscription SubscribeToStreamingNotificationsOnAllFolders(
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "SubscribeToStreamingNotificationsOnAllFolders");

            return this.BuildSubscribeToStreamingNotificationsRequest(null, eventTypes).Execute()[0].Subscription;
        }

        /// <summary>
        /// Begins an asynchronous request to subscribe to streaming notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="eventTypes"></param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSubscribeToStreamingNotificationsOnAllFolders(
            AsyncCallback callback,
            object state,
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "BeginSubscribeToStreamingNotificationsOnAllFolders");

            return this.BuildSubscribeToStreamingNotificationsRequest(null, eventTypes).BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an asynchronous request to subscribe to streaming notifications in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A StreamingSubscription representing the new subscription.</returns>
        public StreamingSubscription EndSubscribeToStreamingNotifications(IAsyncResult asyncResult)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "EndSubscribeToStreamingNotifications");

            var request = AsyncRequestResult.ExtractServiceRequest<SubscribeToStreamingNotificationsRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Subscription;
        }

        /// <summary>
        /// Builds request to subscribe to streaming notifications in the authenticated user's mailbox. 
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A request to subscribe to streaming notifications in the authenticated user's mailbox. </returns>
        private SubscribeToStreamingNotificationsRequest BuildSubscribeToStreamingNotificationsRequest(
            IEnumerable<FolderId> folderIds,
            EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(eventTypes, "eventTypes");

            SubscribeToStreamingNotificationsRequest request = new SubscribeToStreamingNotificationsRequest(this);

            if (folderIds != null)
            {
                request.FolderIds.AddRange(folderIds);
            }

            request.EventTypes.AddRange(eventTypes);

            return request;
        }

        #endregion

        #region Synchronization operations

        /// <summary>
        /// Synchronizes the items of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public ChangeCollection<ItemChange> SyncFolderItems(
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            return this.SyncFolderItems(
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                0, // numberOfDays
                syncScope,
                syncState);
        }

        /// <summary>
        /// Synchronizes the items of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public ChangeCollection<ItemChange> SyncFolderItems(
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            int numberOfDays,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            return this.BuildSyncFolderItemsRequest(
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                numberOfDays,
                syncScope,
                syncState).Execute()[0].Changes;
        }

        /// <summary>
        /// Begins an asynchronous request to synchronize the items of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSyncFolderItems(
            AsyncCallback callback,
            object state,
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            return this.BeginSyncFolderItems(
                callback,
                state,
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                0, // numberOfDays
                syncScope,
                syncState);
        }

        /// <summary>
        /// Begins an asynchronous request to synchronize the items of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSyncFolderItems(
            AsyncCallback callback,
            object state,
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            int numberOfDays,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            return this.BuildSyncFolderItemsRequest(
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                numberOfDays,
                syncScope,
                syncState).BeginExecute(callback, state);
        }

        /// <summary>
        /// Ends an asynchronous request to synchronize the items of a specific folder. 
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public ChangeCollection<ItemChange> EndSyncFolderItems(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<SyncFolderItemsRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Changes;
        }

        /// <summary>
        /// Builds a request to synchronize the items of a specific folder.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A request to synchronize the items of a specific folder.</returns>
        private SyncFolderItemsRequest BuildSyncFolderItemsRequest(
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            int numberOfDays,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            EwsUtilities.ValidateParam(syncFolderId, "syncFolderId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            SyncFolderItemsRequest request = new SyncFolderItemsRequest(this);

            request.SyncFolderId = syncFolderId;
            request.PropertySet = propertySet;
            if (ignoredItemIds != null)
            {
                request.IgnoredItemIds.AddRange(ignoredItemIds);
            }
            request.MaxChangesReturned = maxChangesReturned;
            request.NumberOfDays = numberOfDays;
            request.SyncScope = syncScope;
            request.SyncState = syncState;

            return request;
        }

        /// <summary>
        /// Synchronizes the sub-folders of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with. A null value indicates the root folder of the mailbox.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public ChangeCollection<FolderChange> SyncFolderHierarchy(
            FolderId syncFolderId,
            PropertySet propertySet,
            string syncState)
        {
            return this.BuildSyncFolderHierarchyRequest(
                syncFolderId,
                propertySet,
                syncState).Execute()[0].Changes;
        }

        /// <summary>
        /// Begins an asynchronous request to synchronize the sub-folders of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with. A null value indicates the root folder of the mailbox.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSyncFolderHierarchy(
            AsyncCallback callback,
            object state,
            FolderId syncFolderId,
            PropertySet propertySet,
            string syncState)
        {
            return this.BuildSyncFolderHierarchyRequest(
                syncFolderId,
                propertySet,
                syncState).BeginExecute(callback, state);
        }

        /// <summary>
        /// Synchronizes the entire folder hierarchy of the mailbox this Service is connected to. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public ChangeCollection<FolderChange> SyncFolderHierarchy(PropertySet propertySet, string syncState)
        {
            return this.SyncFolderHierarchy(
                null,
                propertySet,
                syncState);
        }

        /// <summary>
        /// Begins an asynchronous request to synchronize the entire folder hierarchy of the mailbox this Service is connected to. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginSyncFolderHierarchy(
            AsyncCallback callback,
            object state,
            PropertySet propertySet,
            string syncState)
        {
            return this.BeginSyncFolderHierarchy(
                callback,
                state,
                null,
                propertySet,
                syncState);
        }

        /// <summary>
        /// Ends an asynchronous request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public ChangeCollection<FolderChange> EndSyncFolderHierarchy(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<SyncFolderHierarchyRequest>(this, asyncResult);

            return request.EndExecute(asyncResult)[0].Changes;
        }

        /// <summary>
        /// Builds a request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with. A null value indicates the root folder of the mailbox.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.</returns>
        private SyncFolderHierarchyRequest BuildSyncFolderHierarchyRequest(
            FolderId syncFolderId,
            PropertySet propertySet,
            string syncState)
        {
            EwsUtilities.ValidateParamAllowNull(syncFolderId, "syncFolderId");  // Null syncFolderId is allowed
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            SyncFolderHierarchyRequest request = new SyncFolderHierarchyRequest(this);

            request.PropertySet = propertySet;
            request.SyncFolderId = syncFolderId;
            request.SyncState = syncState;

            return request;
        }

        #endregion

        #region Availability operations

        /// <summary>
        /// Gets Out of Office (OOF) settings for a specific user. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the user for which to retrieve OOF settings.</param>
        /// <returns>An OofSettings instance containing OOF information for the specified user.</returns>
        public OofSettings GetUserOofSettings(string smtpAddress)
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");

            GetUserOofSettingsRequest request = new GetUserOofSettingsRequest(this);

            request.SmtpAddress = smtpAddress;

            return request.Execute().OofSettings;
        }

        /// <summary>
        /// Sets the Out of Office (OOF) settings for a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the user for which to set OOF settings.</param>
        /// <param name="oofSettings">The OOF settings.</param>
        public void SetUserOofSettings(string smtpAddress, OofSettings oofSettings)
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");
            EwsUtilities.ValidateParam(oofSettings, "oofSettings");

            SetUserOofSettingsRequest request = new SetUserOofSettingsRequest(this);

            request.SmtpAddress = smtpAddress;
            request.OofSettings = oofSettings;

            request.Execute();
        }

        /// <summary>
        /// Gets detailed information about the availability of a set of users, rooms, and resources within a
        /// specified time window.
        /// </summary>
        /// <param name="attendees">The attendees for which to retrieve availability information.</param>
        /// <param name="timeWindow">The time window in which to retrieve user availability information.</param>
        /// <param name="requestedData">The requested data (free/busy and/or suggestions).</param>
        /// <param name="options">The options controlling the information returned.</param>
        /// <returns>
        /// The availability information for each user appears in a unique FreeBusyResponse object. The order of users
        /// in the request determines the order of availability data for each user in the response.
        /// </returns>
        public GetUserAvailabilityResults GetUserAvailability(
            IEnumerable<AttendeeInfo> attendees,
            TimeWindow timeWindow,
            AvailabilityData requestedData,
            AvailabilityOptions options)
        {
            EwsUtilities.ValidateParamCollection(attendees, "attendees");
            EwsUtilities.ValidateParam(timeWindow, "timeWindow");
            EwsUtilities.ValidateParam(options, "options");

            GetUserAvailabilityRequest request = new GetUserAvailabilityRequest(this);

            request.Attendees = attendees;
            request.TimeWindow = timeWindow;
            request.RequestedData = requestedData;
            request.Options = options;

            return request.Execute();
        }

        /// <summary>
        /// Gets detailed information about the availability of a set of users, rooms, and resources within a
        /// specified time window.
        /// </summary>
        /// <param name="attendees">The attendees for which to retrieve availability information.</param>
        /// <param name="timeWindow">The time window in which to retrieve user availability information.</param>
        /// <param name="requestedData">The requested data (free/busy and/or suggestions).</param>
        /// <returns>
        /// The availability information for each user appears in a unique FreeBusyResponse object. The order of users
        /// in the request determines the order of availability data for each user in the response.
        /// </returns>
        public GetUserAvailabilityResults GetUserAvailability(
            IEnumerable<AttendeeInfo> attendees,
            TimeWindow timeWindow,
            AvailabilityData requestedData)
        {
            return this.GetUserAvailability(
                attendees,
                timeWindow,
                requestedData,
                new AvailabilityOptions());
        }

        /// <summary>
        /// Retrieves a collection of all room lists in the organization.
        /// </summary>
        /// <returns>An EmailAddressCollection containing all the room lists in the organization.</returns>
        public EmailAddressCollection GetRoomLists()
        {
            GetRoomListsRequest request = new GetRoomListsRequest(this);

            return request.Execute().RoomLists;
        }

        /// <summary>
        /// Retrieves a collection of all rooms in the specified room list in the organization.
        /// </summary>
        /// <param name="emailAddress">The e-mail address of the room list.</param>
        /// <returns>A collection of EmailAddress objects representing all the rooms within the specifed room list.</returns>
        public Collection<EmailAddress> GetRooms(EmailAddress emailAddress)
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");

            GetRoomsRequest request = new GetRoomsRequest(this);

            request.RoomList = emailAddress;

            return request.Execute().Rooms;
        }
        #endregion

        #region Conversation
        /// <summary>
        /// Retrieves a collection of all Conversations in the specified Folder.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <returns>Collection of conversations.</returns>
        public ICollection<Conversation> FindConversation(ViewBase view, FolderId folderId)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2010_SP1,
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);

            return request.Execute().Conversations;
        }

        /// <summary>
        /// Retrieves a collection of all Conversations in the specified Folder.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="anchorMailbox">The anchorMailbox Smtp address to route the request directly to group mailbox.</param>
        /// <returns>Collection of conversations.</returns>
        /// <remarks>
        /// This API designed to be used primarily in groups scenarios where we want to set the
        /// anchor mailbox header so that request is routed directly to the group mailbox backend server.
        /// </remarks>
        public Collection<Conversation> FindGroupConversation(
            ViewBase view,
            FolderId folderId,
            string anchorMailbox)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateParam(anchorMailbox, "anchorMailbox");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2015,
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.AnchorMailbox = anchorMailbox;

            return request.Execute().Conversations;
        }

        /// <summary>
        /// Retrieves a collection of all Conversations in the specified Folder.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>Collection of conversations.</returns>
        public ICollection<Conversation> FindConversation(ViewBase view, FolderId folderId, string queryString)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.QueryString = queryString;

            return request.Execute().Conversations;
        }

        /// <summary>
        /// Searches for and retrieves a collection of Conversations in the specified Folder.
        /// Along with conversations, a list of highlight terms are returned.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <returns>FindConversation results.</returns>
        public FindConversationResults FindConversation(ViewBase view, FolderId folderId, string queryString, bool returnHighlightTerms)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParam(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;

            return request.Execute().Results;
        }

        /// <summary>
        /// Searches for and retrieves a collection of Conversations in the specified Folder.
        /// Along with conversations, a list of highlight terms are returned.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <param name="mailboxScope">The mailbox scope to reference.</param>
        /// <returns>FindConversation results.</returns>
        public FindConversationResults FindConversation(ViewBase view, FolderId folderId, string queryString, bool returnHighlightTerms, MailboxSearchLocation? mailboxScope)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParam(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateParam(folderId, "folderId");

            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;
            request.MailboxScope = mailboxScope;

            return request.Execute().Results;
        }

        /// <summary>
        /// Gets the items for a set of conversations.
        /// </summary>
        /// <param name="conversations">Conversations with items to load.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Sort order of conversation tree nodes.</param>
        /// <param name="mailboxScope">The mailbox scope to reference.</param>
        /// <param name="anchorMailbox">The smtpaddress of the mailbox that hosts the conversations</param>
        /// <param name="maxItemsToReturn">Maximum number of items to return.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>GetConversationItems response.</returns>
        internal ServiceResponseCollection<GetConversationItemsResponse> InternalGetConversationItems(
                            IEnumerable<ConversationRequest> conversations,
                            PropertySet propertySet,
                            IEnumerable<FolderId> foldersToIgnore,
                            ConversationSortOrder? sortOrder,
                            MailboxSearchLocation? mailboxScope,
                            int? maxItemsToReturn,
                            string anchorMailbox,
                            ServiceErrorHandling errorHandling)
        {
            EwsUtilities.ValidateParam(conversations, "conversations");
            EwsUtilities.ValidateParam(propertySet, "itemProperties");
            EwsUtilities.ValidateParamAllowNull(foldersToIgnore, "foldersToIgnore");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013,
                                            "GetConversationItems");

            GetConversationItemsRequest request = new GetConversationItemsRequest(this, errorHandling);
            request.ItemProperties = propertySet;
            request.FoldersToIgnore = new FolderIdCollection(foldersToIgnore);
            request.SortOrder = sortOrder;
            request.MailboxScope = mailboxScope;
            request.MaxItemsToReturn = maxItemsToReturn;
            request.AnchorMailbox = anchorMailbox;
            request.Conversations = conversations.ToList();

            return request.Execute();
        }

        /// <summary>
        /// Gets the items for a set of conversations.
        /// </summary>
        /// <param name="conversations">Conversations with items to load.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <returns>GetConversationItems response.</returns>
        public ServiceResponseCollection<GetConversationItemsResponse> GetConversationItems(
                                                IEnumerable<ConversationRequest> conversations,
                                                PropertySet propertySet,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder)
        {
            return this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                null,               /* sortOrder */
                                null,               /* mailboxScope */
                                null,               /* maxItemsToReturn*/
                                null, /* anchorMailbox */
                                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Gets the items for a conversation.
        /// </summary>
        /// <param name="conversationId">The conversation id.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <returns>ConversationResponseType response.</returns>
        public ConversationResponse GetConversationItems(
                                                ConversationId conversationId,
                                                PropertySet propertySet,
                                                string syncState,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder)
        {
            List<ConversationRequest> conversations = new List<ConversationRequest>();
            conversations.Add(new ConversationRequest(conversationId, syncState));

            return this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                sortOrder,
                                null,           /* mailboxScope */
                                null,           /* maxItemsToReturn */
                                null, /* anchorMailbox */
                                ServiceErrorHandling.ThrowOnError)[0].Conversation;
        }

        /// <summary>
        /// Gets the items for a conversation.
        /// </summary>
        /// <param name="conversationId">The conversation id.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <param name="anchorMailbox">The smtp address of the mailbox hosting the conversations</param>
        /// <returns>ConversationResponseType response.</returns>
        /// <remarks>
        /// This API designed to be used primarily in groups scenarios where we want to set the
        /// anchor mailbox header so that request is routed directly to the group mailbox backend server.
        /// </remarks>
        public ConversationResponse GetGroupConversationItems(
                                                ConversationId conversationId,
                                                PropertySet propertySet,
                                                string syncState,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder,
                                                string anchorMailbox)
        {
            EwsUtilities.ValidateParam(anchorMailbox, "anchorMailbox");

            List<ConversationRequest> conversations = new List<ConversationRequest>();
            conversations.Add(new ConversationRequest(conversationId, syncState));

            return this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                sortOrder,
                                null,           /* mailboxScope */
                                null,           /* maxItemsToReturn */
                                anchorMailbox, /* anchorMailbox */
                                ServiceErrorHandling.ThrowOnError)[0].Conversation;
        }

        /// <summary>
        /// Gets the items for a set of conversations.
        /// </summary>
        /// <param name="conversations">Conversations with items to load.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <param name="mailboxScope">The mailbox scope to reference.</param>
        /// <returns>GetConversationItems response.</returns>
        public ServiceResponseCollection<GetConversationItemsResponse> GetConversationItems(
                                                IEnumerable<ConversationRequest> conversations,
                                                PropertySet propertySet,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder,
                                                MailboxSearchLocation? mailboxScope)
        {
            return this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                null,               /* sortOrder */
                                mailboxScope,
                                null,               /* maxItemsToReturn*/
                                null, /* anchorMailbox */
                                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Applies ConversationAction on the specified conversation.
        /// </summary>
        /// <param name="actionType">ConversationAction</param>
        /// <param name="conversationIds">The conversation ids.</param>
        /// <param name="processRightAway">True to process at once . This is blocking
        /// and false to let the Assistant process it in the back ground</param>
        /// <param name="categories">Catgories that need to be stamped can be null or empty</param>
        /// <param name="enableAlwaysDelete">True moves every current and future messages in the conversation
        /// to deleted items folder. False stops the alwasy delete action. This is applicable only if
        /// the action is AlwaysDelete</param>
        /// <param name="destinationFolderId">Applicable if the action is AlwaysMove. This moves every current message and future
        /// message in the conversation to the specified folder. Can be null if tis is then it stops
        /// the always move action</param>
        /// <param name="errorHandlingMode">The error handling mode.</param>
        /// <returns></returns>
        private ServiceResponseCollection<ServiceResponse> ApplyConversationAction(
                ConversationActionType actionType,
                IEnumerable<ConversationId> conversationIds,
                bool processRightAway,
                StringList categories,
                bool enableAlwaysDelete,
                FolderId destinationFolderId,
                ServiceErrorHandling errorHandlingMode)
        {
            EwsUtilities.Assert(
                actionType == ConversationActionType.AlwaysCategorize ||
                actionType == ConversationActionType.AlwaysMove ||
                actionType == ConversationActionType.AlwaysDelete,
                "ApplyConversationAction",
                "Invalid actionType");

            EwsUtilities.ValidateParam(conversationIds, "conversationId");
            EwsUtilities.ValidateMethodVersion(
                                this,
                                ExchangeVersion.Exchange2010_SP1,
                                "ApplyConversationAction");

            ApplyConversationActionRequest request = new ApplyConversationActionRequest(this, errorHandlingMode);

            foreach (var conversationId in conversationIds)
            {
                ConversationAction action = new ConversationAction();

                action.Action = actionType;
                action.ConversationId = conversationId;
                action.ProcessRightAway = processRightAway;
                action.Categories = categories;
                action.EnableAlwaysDelete = enableAlwaysDelete;
                action.DestinationFolderId = destinationFolderId != null ? new FolderIdWrapper(destinationFolderId) : null;

                request.ConversationActions.Add(action);
            }

            return request.Execute();
        }

        /// <summary>
        /// Applies one time conversation action on items in specified folder inside
        /// the conversation.
        /// </summary>
        /// <param name="actionType">The action.</param>
        /// <param name="idTimePairs">The id time pairs.</param>
        /// <param name="contextFolderId">The context folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <param name="deleteType">Type of the delete.</param>
        /// <param name="isRead">The is read.</param>
        /// <param name="retentionPolicyType">Retention policy type.</param>
        /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
        /// <param name="flag">Flag status.</param>
        /// <param name="suppressReadReceipts">Suppress read receipts flag.</param>
        /// <param name="errorHandlingMode">The error handling mode.</param>
        /// <returns></returns>
        private ServiceResponseCollection<ServiceResponse> ApplyConversationOneTimeAction(
            ConversationActionType actionType,
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idTimePairs,
            FolderId contextFolderId,
            FolderId destinationFolderId,
            DeleteMode? deleteType,
            bool? isRead,
            RetentionType? retentionPolicyType,
            Guid? retentionPolicyTagId,
            Flag flag,
            bool? suppressReadReceipts,
            ServiceErrorHandling errorHandlingMode)
        {
            EwsUtilities.Assert(
                actionType == ConversationActionType.Move ||
                actionType == ConversationActionType.Delete ||
                actionType == ConversationActionType.SetReadState ||
                actionType == ConversationActionType.SetRetentionPolicy ||
                actionType == ConversationActionType.Copy ||
                actionType == ConversationActionType.Flag,
                "ApplyConversationOneTimeAction",
                "Invalid actionType");

            EwsUtilities.ValidateParamCollection(idTimePairs, "idTimePairs");
            EwsUtilities.ValidateMethodVersion(
                                this,
                                ExchangeVersion.Exchange2010_SP1,
                                "ApplyConversationAction");

            ApplyConversationActionRequest request = new ApplyConversationActionRequest(this, errorHandlingMode);

            foreach (var idTimePair in idTimePairs)
            {
                ConversationAction action = new ConversationAction();

                action.Action = actionType;
                action.ConversationId = idTimePair.Key;
                action.ContextFolderId = contextFolderId != null ? new FolderIdWrapper(contextFolderId) : null;
                action.DestinationFolderId = destinationFolderId != null ? new FolderIdWrapper(destinationFolderId) : null;
                action.ConversationLastSyncTime = idTimePair.Value;
                action.IsRead = isRead;
                action.DeleteType = deleteType;
                action.RetentionPolicyType = retentionPolicyType;
                action.RetentionPolicyTagId = retentionPolicyTagId;
                action.Flag = flag;
                action.SuppressReadReceipts = suppressReadReceipts;

                request.ConversationActions.Add(action);
            }

            return request.Execute();
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is always categorized.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="categories">The categories that should be stamped on items in the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once enabling this rule and stamping existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> EnableAlwaysCategorizeItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            IEnumerable<String> categories,
            bool processSynchronously)
        {
            EwsUtilities.ValidateParamCollection(categories, "categories");
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysCategorize,
                        conversationId,
                        processSynchronously,
                        new StringList(categories),
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is no longer categorized.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once disabling this rule and removing the categories from existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> DisableAlwaysCategorizeItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            bool processSynchronously)
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysCategorize,
                        conversationId,
                        processSynchronously,
                        null,
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is always moved to Deleted Items folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once enabling this rule and deleting existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> EnableAlwaysDeleteItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            bool processSynchronously)
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysDelete,
                        conversationId,
                        processSynchronously,
                        null,
                        true,
                        null,
                        ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is no longer moved to Deleted Items folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once disabling this rule and restoring the items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> DisableAlwaysDeleteItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            bool processSynchronously)
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysDelete,
                        conversationId,
                        processSynchronously,
                        null,
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is always moved to a specific folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="destinationFolderId">The Id of the folder to which conversation items should be moved.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once enabling this rule and moving existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> EnableAlwaysMoveItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            FolderId destinationFolderId,
            bool processSynchronously)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysMove,
                        conversationId,
                        processSynchronously,
                        null,
                        false,
                        destinationFolderId,
                        ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is no longer moved to a specific folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationIds">The conversation ids.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once disabling this rule is completely done.
        /// If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> DisableAlwaysMoveItemsInConversations(
            IEnumerable<ConversationId> conversationIds,
            bool processSynchronously)
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysMove,
                        conversationIds,
                        processSynchronously,
                        null,
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Moves the items in the specified conversation to the specified destination folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should be moved and the dateTime conversation was last synced
        /// (Items received after that dateTime will not be moved).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="destinationFolderId">The Id of the destination folder.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> MoveItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Move,
                idLastSyncTimePairs,
                contextFolderId,
                destinationFolderId,
                null,
                null,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Copies the items in the specified conversation to the specified destination folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should be copied and the date and time conversation was last synced
        /// (Items received after that date will not be copied).</param>
        /// <param name="contextFolderId">The context folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> CopyItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            FolderId destinationFolderId)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Copy,
                idLastSyncTimePairs,
                contextFolderId,
                destinationFolderId,
                null,
                null,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Deletes the items in the specified conversation. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should be deleted and the date and time conversation was last synced
        /// (Items received after that date will not be deleted).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <returns></returns>
        public ServiceResponseCollection<ServiceResponse> DeleteItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            DeleteMode deleteMode)
        {
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Delete,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                deleteMode,
                null,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets the read state for items in conversation. Calling this method would
        /// result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their read state set and the date and time conversation
        /// was last synced (Items received after that date will not have their read
        /// state set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="isRead">if set to <c>true</c>, conversation items are marked as read; otherwise they are marked as unread.</param>
        public ServiceResponseCollection<ServiceResponse> SetReadStateForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            bool isRead)
        {
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.SetReadState,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                isRead,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets the read state for items in conversation. Calling this method would
        /// result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their read state set and the date and time conversation
        /// was last synced (Items received after that date will not have their read
        /// state set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="isRead">if set to <c>true</c>, conversation items are marked as read; otherwise they are marked as unread.</param>
        /// <param name="suppressReadReceipts">if set to <c>true</c> read receipts are suppressed.</param>
        public ServiceResponseCollection<ServiceResponse> SetReadStateForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            bool isRead,
            bool suppressReadReceipts)
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetReadStateForItemsInConversations");

            return this.ApplyConversationOneTimeAction(
                ConversationActionType.SetReadState,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                isRead,
                null,
                null,
                null,
                suppressReadReceipts,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets the retention policy for items in conversation. Calling this method would
        /// result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their retention policy set and the date and time conversation
        /// was last synced (Items received after that date will not have their retention
        /// policy set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="retentionPolicyType">Retention policy type.</param>
        /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
        public ServiceResponseCollection<ServiceResponse> SetRetentionPolicyForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            RetentionType retentionPolicyType,
            Guid? retentionPolicyTagId)
        {
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.SetRetentionPolicy,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                null,
                retentionPolicyType,
                retentionPolicyTagId,
                null,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Sets flag status for items in conversation. Calling this method would result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their read state set and the date and time conversation
        /// was last synced (Items received after that date will not have their read
        /// state set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="flagStatus">Flag status to apply to conversation items.</param>
        public ServiceResponseCollection<ServiceResponse> SetFlagStatusForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            Flag flagStatus)
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetFlagStatusForItemsInConversations");

            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Flag,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                null,
                null,
                null,
                flagStatus,
                null,
                ServiceErrorHandling.ReturnErrors);
        }

        #endregion

        #region Id conversion operations

        /// <summary>
        /// Converts multiple Ids from one format to another in a single call to EWS.
        /// </summary>
        /// <param name="ids">The Ids to convert.</param>
        /// <param name="destinationFormat">The destination format.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing conversion results for each specified Ids.</returns>
        private ServiceResponseCollection<ConvertIdResponse> InternalConvertIds(
            IEnumerable<AlternateIdBase> ids,
            IdFormat destinationFormat,
            ServiceErrorHandling errorHandling)
        {
            EwsUtilities.ValidateParamCollection(ids, "ids");

            ConvertIdRequest request = new ConvertIdRequest(this, errorHandling);

            request.Ids.AddRange(ids);
            request.DestinationFormat = destinationFormat;

            return request.Execute();
        }

        /// <summary>
        /// Converts multiple Ids from one format to another in a single call to EWS.
        /// </summary>
        /// <param name="ids">The Ids to convert.</param>
        /// <param name="destinationFormat">The destination format.</param>
        /// <returns>A ServiceResponseCollection providing conversion results for each specified Ids.</returns>
        public ServiceResponseCollection<ConvertIdResponse> ConvertIds(IEnumerable<AlternateIdBase> ids, IdFormat destinationFormat)
        {
            EwsUtilities.ValidateParamCollection(ids, "ids");

            return this.InternalConvertIds(
                ids,
                destinationFormat,
                ServiceErrorHandling.ReturnErrors);
        }

        /// <summary>
        /// Converts Id from one format to another in a single call to EWS.
        /// </summary>
        /// <param name="id">The Id to convert.</param>
        /// <param name="destinationFormat">The destination format.</param>
        /// <returns>The converted Id.</returns>
        public AlternateIdBase ConvertId(AlternateIdBase id, IdFormat destinationFormat)
        {
            EwsUtilities.ValidateParam(id, "id");

            ServiceResponseCollection<ConvertIdResponse> responses = this.InternalConvertIds(
                new AlternateIdBase[] { id },
                destinationFormat,
                ServiceErrorHandling.ThrowOnError);

            return responses[0].ConvertedId;
        }

        #endregion

        #region Delegate management operations

        /// <summary>
        /// Adds delegates to a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to add delegates to.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to add.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Collection<DelegateUserResponse> AddDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            params DelegateUser[] delegateUsers)
        {
            return this.AddDelegates(
                mailbox,
                meetingRequestsDeliveryScope,
                (IEnumerable<DelegateUser>)delegateUsers);
        }

        /// <summary>
        /// Adds delegates to a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to add delegates to.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to add.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Collection<DelegateUserResponse> AddDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            IEnumerable<DelegateUser> delegateUsers)
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");
            EwsUtilities.ValidateParamCollection(delegateUsers, "delegateUsers");

            AddDelegateRequest request = new AddDelegateRequest(this);

            request.Mailbox = mailbox;
            request.DelegateUsers.AddRange(delegateUsers);
            request.MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope;

            DelegateManagementResponse response = request.Execute();
            return response.DelegateUserResponses;
        }

        /// <summary>
        /// Updates delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to update delegates on.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to update.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Collection<DelegateUserResponse> UpdateDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            params DelegateUser[] delegateUsers)
        {
            return this.UpdateDelegates(
                mailbox,
                meetingRequestsDeliveryScope,
                (IEnumerable<DelegateUser>)delegateUsers);
        }

        /// <summary>
        /// Updates delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to update delegates on.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to update.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Collection<DelegateUserResponse> UpdateDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            IEnumerable<DelegateUser> delegateUsers)
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");
            EwsUtilities.ValidateParamCollection(delegateUsers, "delegateUsers");

            UpdateDelegateRequest request = new UpdateDelegateRequest(this);

            request.Mailbox = mailbox;
            request.DelegateUsers.AddRange(delegateUsers);
            request.MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope;

            DelegateManagementResponse response = request.Execute();
            return response.DelegateUserResponses;
        }

        /// <summary>
        /// Removes delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to remove delegates from.</param>
        /// <param name="userIds">The Ids of the delegate users to remove.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Collection<DelegateUserResponse> RemoveDelegates(Mailbox mailbox, params UserId[] userIds)
        {
            return this.RemoveDelegates(mailbox, (IEnumerable<UserId>)userIds);
        }

        /// <summary>
        /// Removes delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to remove delegates from.</param>
        /// <param name="userIds">The Ids of the delegate users to remove.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Collection<DelegateUserResponse> RemoveDelegates(Mailbox mailbox, IEnumerable<UserId> userIds)
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");
            EwsUtilities.ValidateParamCollection(userIds, "userIds");

            RemoveDelegateRequest request = new RemoveDelegateRequest(this);

            request.Mailbox = mailbox;
            request.UserIds.AddRange(userIds);

            DelegateManagementResponse response = request.Execute();
            return response.DelegateUserResponses;
        }

        /// <summary>
        /// Retrieves the delegates of a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to retrieve the delegates of.</param>
        /// <param name="includePermissions">Indicates whether detailed permissions should be returned fro each delegate.</param>
        /// <param name="userIds">The optional Ids of the delegate users to retrieve.</param>
        /// <returns>A GetDelegateResponse providing the results of the operation.</returns>
        public DelegateInformation GetDelegates(
            Mailbox mailbox,
            bool includePermissions,
            params UserId[] userIds)
        {
            return this.GetDelegates(
                mailbox,
                includePermissions,
                (IEnumerable<UserId>)userIds);
        }

        /// <summary>
        /// Retrieves the delegates of a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to retrieve the delegates of.</param>
        /// <param name="includePermissions">Indicates whether detailed permissions should be returned fro each delegate.</param>
        /// <param name="userIds">The optional Ids of the delegate users to retrieve.</param>
        /// <returns>A GetDelegateResponse providing the results of the operation.</returns>
        public DelegateInformation GetDelegates(
            Mailbox mailbox,
            bool includePermissions,
            IEnumerable<UserId> userIds)
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");

            GetDelegateRequest request = new GetDelegateRequest(this);

            request.Mailbox = mailbox;
            request.UserIds.AddRange(userIds);
            request.IncludePermissions = includePermissions;

            GetDelegateResponse response = request.Execute();
            DelegateInformation delegateInformation = new DelegateInformation(
                response.DelegateUserResponses,
                response.MeetingRequestsDeliveryScope);

            return delegateInformation;
        }

        #endregion

        #region UserConfiguration operations

        /// <summary>
        /// Creates a UserConfiguration.
        /// </summary>
        /// <param name="userConfiguration">The UserConfiguration.</param>
        internal void CreateUserConfiguration(UserConfiguration userConfiguration)
        {
            EwsUtilities.ValidateParam(userConfiguration, "userConfiguration");

            CreateUserConfigurationRequest request = new CreateUserConfigurationRequest(this);

            request.UserConfiguration = userConfiguration;

            request.Execute();
        }

        /// <summary>
        /// Deletes a UserConfiguration.
        /// </summary>
        /// <param name="name">Name of the UserConfiguration to retrieve.</param>
        /// <param name="parentFolderId">Id of the folder containing the UserConfiguration.</param>
        internal void DeleteUserConfiguration(string name, FolderId parentFolderId)
        {
            EwsUtilities.ValidateParam(name, "name");
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");

            DeleteUserConfigurationRequest request = new DeleteUserConfigurationRequest(this);

            request.Name = name;
            request.ParentFolderId = parentFolderId;

            request.Execute();
        }

        /// <summary>
        /// Gets a UserConfiguration.
        /// </summary>
        /// <param name="name">Name of the UserConfiguration to retrieve.</param>
        /// <param name="parentFolderId">Id of the folder containing the UserConfiguration.</param>
        /// <param name="properties">Properties to retrieve.</param>
        /// <returns>A UserConfiguration.</returns>
        internal UserConfiguration GetUserConfiguration(
            string name,
            FolderId parentFolderId,
            UserConfigurationProperties properties)
        {
            EwsUtilities.ValidateParam(name, "name");
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");

            GetUserConfigurationRequest request = new GetUserConfigurationRequest(this);

            request.Name = name;
            request.ParentFolderId = parentFolderId;
            request.Properties = properties;

            return request.Execute()[0].UserConfiguration;
        }

        /// <summary>
        /// Loads the properties of the specified userConfiguration.
        /// </summary>
        /// <param name="userConfiguration">The userConfiguration containing properties to load.</param>
        /// <param name="properties">Properties to retrieve.</param>
        internal void LoadPropertiesForUserConfiguration(UserConfiguration userConfiguration, UserConfigurationProperties properties)
        {
            EwsUtilities.Assert(
                userConfiguration != null,
                "ExchangeService.LoadPropertiesForUserConfiguration",
                "userConfiguration is null");

            GetUserConfigurationRequest request = new GetUserConfigurationRequest(this);

            request.UserConfiguration = userConfiguration;
            request.Properties = properties;

            request.Execute();
        }

        /// <summary>
        /// Updates a UserConfiguration.
        /// </summary>
        /// <param name="userConfiguration">The UserConfiguration.</param>
        internal void UpdateUserConfiguration(UserConfiguration userConfiguration)
        {
            EwsUtilities.ValidateParam(userConfiguration, "userConfiguration");

            UpdateUserConfigurationRequest request = new UpdateUserConfigurationRequest(this);

            request.UserConfiguration = userConfiguration;

            request.Execute();
        }

        #endregion

        #region InboxRule operations
        /// <summary>
        /// Retrieves inbox rules of the authenticated user.
        /// </summary>
        /// <returns>A RuleCollection object containing the authenticated user's inbox rules.</returns>
        public RuleCollection GetInboxRules()
        {
            GetInboxRulesRequest request = new GetInboxRulesRequest(this);

            return request.Execute().Rules;
        }

        /// <summary>
        /// Retrieves the inbox rules of the specified user.
        /// </summary>
        /// <param name="mailboxSmtpAddress">The SMTP address of the user whose inbox rules should be retrieved.</param>
        /// <returns>A RuleCollection object containing the inbox rules of the specified user.</returns>
        public RuleCollection GetInboxRules(string mailboxSmtpAddress)
        {
            EwsUtilities.ValidateParam(mailboxSmtpAddress, "MailboxSmtpAddress");

            GetInboxRulesRequest request = new GetInboxRulesRequest(this);
            request.MailboxSmtpAddress = mailboxSmtpAddress;

            return request.Execute().Rules;
        }

        /// <summary>
        /// Updates the authenticated user's inbox rules by applying the specified operations.
        /// </summary>
        /// <param name="operations">The operations that should be applied to the user's inbox rules.</param>
        /// <param name="removeOutlookRuleBlob">Indicate whether or not to remove Outlook Rule Blob.</param>
        public void UpdateInboxRules(
            IEnumerable<RuleOperation> operations,
            bool removeOutlookRuleBlob)
        {
            UpdateInboxRulesRequest request = new UpdateInboxRulesRequest(this);
            request.InboxRuleOperations = operations;
            request.RemoveOutlookRuleBlob = removeOutlookRuleBlob;
            request.Execute();
        }

        /// <summary>
        /// Update the specified user's inbox rules by applying the specified operations.
        /// </summary>
        /// <param name="operations">The operations that should be applied to the user's inbox rules.</param>
        /// <param name="removeOutlookRuleBlob">Indicate whether or not to remove Outlook Rule Blob.</param>
        /// <param name="mailboxSmtpAddress">The SMTP address of the user whose inbox rules should be updated.</param>
        public void UpdateInboxRules(
            IEnumerable<RuleOperation> operations,
            bool removeOutlookRuleBlob,
            string mailboxSmtpAddress)
        {
            UpdateInboxRulesRequest request = new UpdateInboxRulesRequest(this);
            request.InboxRuleOperations = operations;
            request.RemoveOutlookRuleBlob = removeOutlookRuleBlob;
            request.MailboxSmtpAddress = mailboxSmtpAddress;
            request.Execute();
        }
        #endregion

        #region eDiscovery/Compliance operations

        /// <summary>
        /// Get discovery search configuration
        /// </summary>
        /// <param name="searchId">Search Id</param>
        /// <param name="expandGroupMembership">True if want to expand group membership</param>
        /// <param name="inPlaceHoldConfigurationOnly">True if only want the inplacehold configuration</param>
        /// <returns>Service response object</returns>
        public GetDiscoverySearchConfigurationResponse GetDiscoverySearchConfiguration(string searchId, bool expandGroupMembership, bool inPlaceHoldConfigurationOnly)
        {
            GetDiscoverySearchConfigurationRequest request = new GetDiscoverySearchConfigurationRequest(this);
            request.SearchId = searchId;
            request.ExpandGroupMembership = expandGroupMembership;
            request.InPlaceHoldConfigurationOnly = inPlaceHoldConfigurationOnly;

            return request.Execute();
        }

        /// <summary>
        /// Get searchable mailboxes
        /// </summary>
        /// <param name="searchFilter">Search filter</param>
        /// <param name="expandGroupMembership">True if want to expand group membership</param>
        /// <returns>Service response object</returns>
        public GetSearchableMailboxesResponse GetSearchableMailboxes(string searchFilter, bool expandGroupMembership)
        {
            GetSearchableMailboxesRequest request = new GetSearchableMailboxesRequest(this);
            request.SearchFilter = searchFilter;
            request.ExpandGroupMembership = expandGroupMembership;

            return request.Execute();
        }

        /// <summary>
        /// Search mailboxes
        /// </summary>
        /// <param name="mailboxQueries">Collection of query and mailboxes</param>
        /// <param name="resultType">Search result type</param>
        /// <returns>Collection of search mailboxes response object</returns>
        public ServiceResponseCollection<SearchMailboxesResponse> SearchMailboxes(IEnumerable<MailboxQuery> mailboxQueries, SearchResultType resultType)
        {
            SearchMailboxesRequest request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors);
            if (mailboxQueries != null)
            {
                request.SearchQueries.AddRange(mailboxQueries);
            }

            request.ResultType = resultType;

            return request.Execute();
        }

        /// <summary>
        /// Search mailboxes
        /// </summary>
        /// <param name="mailboxQueries">Collection of query and mailboxes</param>
        /// <param name="resultType">Search result type</param>
        /// <param name="sortByProperty">Sort by property name</param>
        /// <param name="sortOrder">Sort order</param>
        /// <param name="pageSize">Page size</param>
        /// <param name="pageDirection">Page navigation direction</param>
        /// <param name="pageItemReference">Item reference used for paging</param>
        /// <returns>Collection of search mailboxes response object</returns>
        public ServiceResponseCollection<SearchMailboxesResponse> SearchMailboxes(
            IEnumerable<MailboxQuery> mailboxQueries,
            SearchResultType resultType,
            string sortByProperty,
            SortDirection sortOrder,
            int pageSize,
            SearchPageDirection pageDirection,
            string pageItemReference)
        {
            SearchMailboxesRequest request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors);
            if (mailboxQueries != null)
            {
                request.SearchQueries.AddRange(mailboxQueries);
            }

            request.ResultType = resultType;
            request.SortByProperty = sortByProperty;
            request.SortOrder = sortOrder;
            request.PageSize = pageSize;
            request.PageDirection = pageDirection;
            request.PageItemReference = pageItemReference;

            return request.Execute();
        }

        /// <summary>
        /// Search mailboxes
        /// </summary>
        /// <param name="searchParameters">Search mailboxes parameters</param>
        /// <returns>Collection of search mailboxes response object</returns>
        public ServiceResponseCollection<SearchMailboxesResponse> SearchMailboxes(SearchMailboxesParameters searchParameters)
        {
            EwsUtilities.ValidateParam(searchParameters, "searchParameters");
            EwsUtilities.ValidateParam(searchParameters.SearchQueries, "searchParameters.SearchQueries");

            SearchMailboxesRequest request = this.CreateSearchMailboxesRequest(searchParameters);
            return request.Execute();
        }

        /// <summary>
        /// Asynchronous call to search mailboxes
        /// </summary>
        /// <param name="callback">callback</param>
        /// <param name="state">state</param>
        /// <param name="searchParameters">search parameters</param>
        /// <returns>Async result</returns>
        public IAsyncResult BeginSearchMailboxes(
            AsyncCallback callback,
            object state,
            SearchMailboxesParameters searchParameters)
        {
            EwsUtilities.ValidateParam(searchParameters, "searchParameters");
            EwsUtilities.ValidateParam(searchParameters.SearchQueries, "searchParameters.SearchQueries");

            SearchMailboxesRequest request = this.CreateSearchMailboxesRequest(searchParameters);
            return request.BeginExecute(callback, state);
        }

        /// <summary>
        /// Asynchronous call to end search mailboxes
        /// </summary>
        /// <param name="asyncResult"></param>
        /// <returns></returns>
        public ServiceResponseCollection<SearchMailboxesResponse> EndSearchMailboxes(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<SearchMailboxesRequest>(this, asyncResult);

            return request.EndExecute(asyncResult);
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <param name="actionType">Action type</param>
        /// <param name="query">Query string</param>
        /// <param name="mailboxes">Collection of mailboxes</param>
        /// <returns>Service response object</returns>
        public SetHoldOnMailboxesResponse SetHoldOnMailboxes(string holdId, HoldAction actionType, string query, string[] mailboxes)
        {
            SetHoldOnMailboxesRequest request = new SetHoldOnMailboxesRequest(this);
            request.HoldId = holdId;
            request.ActionType = actionType;
            request.Query = query;
            request.Mailboxes = mailboxes;
            request.InPlaceHoldIdentity = null;

            return request.Execute();
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <param name="actionType">Action type</param>
        /// <param name="query">Query string</param>
        /// <param name="inPlaceHoldIdentity">in-place hold identity</param>
        /// <returns>Service response object</returns>
        public SetHoldOnMailboxesResponse SetHoldOnMailboxes(string holdId, HoldAction actionType, string query, string inPlaceHoldIdentity)
        {
            return this.SetHoldOnMailboxes(holdId, actionType, query, inPlaceHoldIdentity, null);
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <param name="actionType">Action type</param>
        /// <param name="query">Query string</param>
        /// <param name="inPlaceHoldIdentity">in-place hold identity</param>
        /// <param name="itemHoldPeriod">item hold period</param>
        /// <returns>Service response object</returns>
        public SetHoldOnMailboxesResponse SetHoldOnMailboxes(string holdId, HoldAction actionType, string query, string inPlaceHoldIdentity, string itemHoldPeriod)
        {
            SetHoldOnMailboxesRequest request = new SetHoldOnMailboxesRequest(this);
            request.HoldId = holdId;
            request.ActionType = actionType;
            request.Query = query;
            request.Mailboxes = null;
            request.InPlaceHoldIdentity = inPlaceHoldIdentity;
            request.ItemHoldPeriod = itemHoldPeriod;

            return request.Execute();
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="parameters">Set hold parameters</param>
        /// <returns>Service response object</returns>
        public SetHoldOnMailboxesResponse SetHoldOnMailboxes(SetHoldOnMailboxesParameters parameters)
        {
            EwsUtilities.ValidateParam(parameters, "parameters");

            SetHoldOnMailboxesRequest request = new SetHoldOnMailboxesRequest(this);
            request.HoldId = parameters.HoldId;
            request.ActionType = parameters.ActionType;
            request.Query = parameters.Query;
            request.Mailboxes = parameters.Mailboxes;
            request.Language = parameters.Language;
            request.InPlaceHoldIdentity = parameters.InPlaceHoldIdentity;

            return request.Execute();
        }

        /// <summary>
        /// Get hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <returns>Service response object</returns>
        public GetHoldOnMailboxesResponse GetHoldOnMailboxes(string holdId)
        {
            GetHoldOnMailboxesRequest request = new GetHoldOnMailboxesRequest(this);
            request.HoldId = holdId;

            return request.Execute();
        }

        /// <summary>
        /// Get non indexable item details
        /// </summary>
        /// <param name="mailboxes">Array of mailbox legacy DN</param>
        /// <returns>Service response object</returns>
        public GetNonIndexableItemDetailsResponse GetNonIndexableItemDetails(string[] mailboxes)
        {
            return this.GetNonIndexableItemDetails(mailboxes, null, null, null);
        }

        /// <summary>
        /// Get non indexable item details
        /// </summary>
        /// <param name="mailboxes">Array of mailbox legacy DN</param>
        /// <param name="pageSize">The page size</param>
        /// <param name="pageItemReference">Page item reference</param>
        /// <param name="pageDirection">Page direction</param>
        /// <returns>Service response object</returns>
        public GetNonIndexableItemDetailsResponse GetNonIndexableItemDetails(string[] mailboxes, int? pageSize, string pageItemReference, SearchPageDirection? pageDirection)
        {
            GetNonIndexableItemDetailsParameters parameters = new GetNonIndexableItemDetailsParameters
            {
                Mailboxes = mailboxes,
                PageSize = pageSize,
                PageItemReference = pageItemReference,
                PageDirection = pageDirection,
                SearchArchiveOnly = false,
            };

            return GetNonIndexableItemDetails(parameters);
        }

        /// <summary>
        /// Get non indexable item details
        /// </summary>
        /// <param name="parameters">Get non indexable item details parameters</param>
        /// <returns>Service response object</returns>
        public GetNonIndexableItemDetailsResponse GetNonIndexableItemDetails(GetNonIndexableItemDetailsParameters parameters)
        {
            GetNonIndexableItemDetailsRequest request = this.CreateGetNonIndexableItemDetailsRequest(parameters);

            return request.Execute();
        }

        /// <summary>
        /// Asynchronous call to get non indexable item details
        /// </summary>
        /// <param name="callback">callback</param>
        /// <param name="state">state</param>
        /// <param name="parameters">Get non indexable item details parameters</param>
        /// <returns>Async result</returns>
        public IAsyncResult BeginGetNonIndexableItemDetails(
            AsyncCallback callback,
            object state,
            GetNonIndexableItemDetailsParameters parameters)
        {
            GetNonIndexableItemDetailsRequest request = this.CreateGetNonIndexableItemDetailsRequest(parameters);
            return request.BeginExecute(callback, state);
        }

        /// <summary>
        /// Asynchronous call to get non indexable item details
        /// </summary>
        /// <param name="asyncResult"></param>
        /// <returns></returns>
        public GetNonIndexableItemDetailsResponse EndGetNonIndexableItemDetails(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<GetNonIndexableItemDetailsRequest>(this, asyncResult);

            return (GetNonIndexableItemDetailsResponse)request.EndInternalExecute(asyncResult);
        }

        /// <summary>
        /// Get non indexable item statistics
        /// </summary>
        /// <param name="mailboxes">Array of mailbox legacy DN</param>
        /// <returns>Service response object</returns>
        public GetNonIndexableItemStatisticsResponse GetNonIndexableItemStatistics(string[] mailboxes)
        {
            GetNonIndexableItemStatisticsParameters parameters = new GetNonIndexableItemStatisticsParameters
            {
                Mailboxes = mailboxes,
                SearchArchiveOnly = false,
            };

            return this.GetNonIndexableItemStatistics(parameters);
        }

        /// <summary>
        /// Get non indexable item statistics
        /// </summary>
        /// <param name="parameters">Get non indexable item statistics parameters</param>
        /// <returns>Service response object</returns>
        public GetNonIndexableItemStatisticsResponse GetNonIndexableItemStatistics(GetNonIndexableItemStatisticsParameters parameters)
        {
            GetNonIndexableItemStatisticsRequest request = this.CreateGetNonIndexableItemStatisticsRequest(parameters);

            return request.Execute();
        }

        /// <summary>
        /// Asynchronous call to get non indexable item statistics
        /// </summary>
        /// <param name="callback">callback</param>
        /// <param name="state">state</param>
        /// <param name="parameters">Get non indexable item statistics parameters</param>
        /// <returns>Async result</returns>
        public IAsyncResult BeginGetNonIndexableItemStatistics(
            AsyncCallback callback,
            object state,
            GetNonIndexableItemStatisticsParameters parameters)
        {
            GetNonIndexableItemStatisticsRequest request = this.CreateGetNonIndexableItemStatisticsRequest(parameters);
            return request.BeginExecute(callback, state);
        }

        /// <summary>
        /// Asynchronous call to get non indexable item statistics
        /// </summary>
        /// <param name="asyncResult"></param>
        /// <returns></returns>
        public GetNonIndexableItemStatisticsResponse EndGetNonIndexableItemStatistics(IAsyncResult asyncResult)
        {
            var request = AsyncRequestResult.ExtractServiceRequest<GetNonIndexableItemStatisticsRequest>(this, asyncResult);

            return (GetNonIndexableItemStatisticsResponse)request.EndInternalExecute(asyncResult);
        }

        /// <summary>
        /// Create get non indexable item details request
        /// </summary>
        /// <param name="parameters">Get non indexable item details parameters</param>
        /// <returns>GetNonIndexableItemDetails request</returns>
        private GetNonIndexableItemDetailsRequest CreateGetNonIndexableItemDetailsRequest(GetNonIndexableItemDetailsParameters parameters)
        {
            EwsUtilities.ValidateParam(parameters, "parameters");
            EwsUtilities.ValidateParam(parameters.Mailboxes, "parameters.Mailboxes");

            GetNonIndexableItemDetailsRequest request = new GetNonIndexableItemDetailsRequest(this);
            request.Mailboxes = parameters.Mailboxes;
            request.PageSize = parameters.PageSize;
            request.PageItemReference = parameters.PageItemReference;
            request.PageDirection = parameters.PageDirection;
            request.SearchArchiveOnly = parameters.SearchArchiveOnly;

            return request;
        }

        /// <summary>
        /// Create get non indexable item statistics request
        /// </summary>
        /// <param name="parameters">Get non indexable item statistics parameters</param>
        /// <returns>Service response object</returns>
        private GetNonIndexableItemStatisticsRequest CreateGetNonIndexableItemStatisticsRequest(GetNonIndexableItemStatisticsParameters parameters)
        {
            EwsUtilities.ValidateParam(parameters, "parameters");
            EwsUtilities.ValidateParam(parameters.Mailboxes, "parameters.Mailboxes");

            GetNonIndexableItemStatisticsRequest request = new GetNonIndexableItemStatisticsRequest(this);
            request.Mailboxes = parameters.Mailboxes;
            request.SearchArchiveOnly = parameters.SearchArchiveOnly;

            return request;
        }

        /// <summary>
        /// Creates SearchMailboxesRequest from SearchMailboxesParameters
        /// </summary>
        /// <param name="searchParameters">search parameters</param>
        /// <returns>request object</returns>
        private SearchMailboxesRequest CreateSearchMailboxesRequest(SearchMailboxesParameters searchParameters)
        {
            SearchMailboxesRequest request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors);
            request.SearchQueries.AddRange(searchParameters.SearchQueries);
            request.ResultType = searchParameters.ResultType;
            request.PreviewItemResponseShape = searchParameters.PreviewItemResponseShape;
            request.SortByProperty = searchParameters.SortBy;
            request.SortOrder = searchParameters.SortOrder;
            request.Language = searchParameters.Language;
            request.PerformDeduplication = searchParameters.PerformDeduplication;
            request.PageSize = searchParameters.PageSize;
            request.PageDirection = searchParameters.PageDirection;
            request.PageItemReference = searchParameters.PageItemReference;

            return request;
        }
        #endregion

        #region MRM operations

        /// <summary>
        /// Get user retention policy tags.
        /// </summary>
        /// <returns>Service response object.</returns>
        public GetUserRetentionPolicyTagsResponse GetUserRetentionPolicyTags()
        {
            GetUserRetentionPolicyTagsRequest request = new GetUserRetentionPolicyTagsRequest(this);

            return request.Execute();
        }

        #endregion

        #region Autodiscover

        /// <summary>
        /// Default implementation of AutodiscoverRedirectionUrlValidationCallback.
        /// Always returns true indicating that the URL can be used.
        /// </summary>
        /// <param name="redirectionUrl">The redirection URL.</param>
        /// <returns>Returns true.</returns>
        private bool DefaultAutodiscoverRedirectionUrlValidationCallback(string redirectionUrl)
        {
            throw new AutodiscoverLocalException(string.Format(Strings.AutodiscoverRedirectBlocked, redirectionUrl));
        }

        /// <summary>
        /// Initializes the Url property to the Exchange Web Services URL for the specified e-mail address by
        /// calling the Autodiscover service.
        /// </summary>
        /// <param name="emailAddress">The email address to use.</param>
        public void AutodiscoverUrl(string emailAddress)
        {
            this.AutodiscoverUrl(emailAddress, this.DefaultAutodiscoverRedirectionUrlValidationCallback);
        }

        /// <summary>
        /// Initializes the Url property to the Exchange Web Services URL for the specified e-mail address by
        /// calling the Autodiscover service.
        /// </summary>
        /// <param name="emailAddress">The email address to use.</param>
        /// <param name="validateRedirectionUrlCallback">The callback used to validate redirection URL.</param>
        public void AutodiscoverUrl(string emailAddress, AutodiscoverRedirectionUrlValidationCallback validateRedirectionUrlCallback)
        {
            Uri exchangeServiceUrl;

            if (this.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                try
                {
                    exchangeServiceUrl = this.GetAutodiscoverUrl(
                        emailAddress,
                        this.RequestedServerVersion,
                        validateRedirectionUrlCallback);

                    this.Url = this.AdjustServiceUriFromCredentials(exchangeServiceUrl);
                    return;
                }
                catch (AutodiscoverLocalException ex)
                {
                    this.TraceMessage(
                        TraceFlags.AutodiscoverResponse,
                        string.Format("Autodiscover service call failed with error '{0}'. Will try legacy service", ex.Message));
                }
                catch (ServiceRemoteException ex)
                {
                    // Special case: if the caller's account is locked we want to return this exception, not continue.
                    if (ex is AccountIsLockedException)
                    {
                        throw;
                    }

                    this.TraceMessage(
                        TraceFlags.AutodiscoverResponse,
                        string.Format("Autodiscover service call failed with error '{0}'. Will try legacy service", ex.Message));
                }
            }

            // Try legacy Autodiscover provider
            exchangeServiceUrl = this.GetAutodiscoverUrl(
                emailAddress,
                ExchangeVersion.Exchange2007_SP1,
                validateRedirectionUrlCallback);

            this.Url = this.AdjustServiceUriFromCredentials(exchangeServiceUrl);
        }

        /// <summary>
        /// Adjusts the service URI based on the current type of credentials.
        /// </summary>
        /// <remarks>
        /// Autodiscover will always return the "plain" EWS endpoint URL but if the client
        /// is using WindowsLive credentials, ExchangeService needs to use the WS-Security endpoint.
        /// </remarks>
        /// <param name="uri">The URI.</param>
        /// <returns>Adjusted URL.</returns>
        private Uri AdjustServiceUriFromCredentials(Uri uri)
        {
            return (this.Credentials != null)
                ? this.Credentials.AdjustUrl(uri)
                : uri;
        }

        /// <summary>
        /// Gets the EWS URL from Autodiscover.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        /// <param name="requestedServerVersion">Exchange version.</param>
        /// <param name="validateRedirectionUrlCallback">The validate redirection URL callback.</param>
        /// <returns>Ews URL</returns>
        private Uri GetAutodiscoverUrl(
            string emailAddress,
            ExchangeVersion requestedServerVersion,
            AutodiscoverRedirectionUrlValidationCallback validateRedirectionUrlCallback)
        {
            AutodiscoverService autodiscoverService = new AutodiscoverService(this, requestedServerVersion)
            {
                RedirectionUrlValidationCallback = validateRedirectionUrlCallback,
                EnableScpLookup = this.EnableScpLookup
            };

            GetUserSettingsResponse response = autodiscoverService.GetUserSettings(
                emailAddress,
                UserSettingName.InternalEwsUrl,
                UserSettingName.ExternalEwsUrl);

            switch (response.ErrorCode)
            {
                case AutodiscoverErrorCode.NoError:
                    return this.GetEwsUrlFromResponse(response, autodiscoverService.IsExternal.GetValueOrDefault(true));

                case AutodiscoverErrorCode.InvalidUser:
                    throw new ServiceRemoteException(
                        string.Format(Strings.InvalidUser, emailAddress));

                case AutodiscoverErrorCode.InvalidRequest:
                    throw new ServiceRemoteException(
                        string.Format(Strings.InvalidAutodiscoverRequest, response.ErrorMessage));

                default:
                    this.TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("No EWS Url returned for user {0}, error code is {1}", emailAddress, response.ErrorCode));

                    throw new ServiceRemoteException(response.ErrorMessage);
            }
        }

        /// <summary>
        /// Gets the EWS URL from Autodiscover GetUserSettings response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="isExternal">If true, Autodiscover call was made externally.</param>
        /// <returns>EWS URL.</returns>
        private Uri GetEwsUrlFromResponse(GetUserSettingsResponse response, bool isExternal)
        {
            string uriString;

            // Figure out which URL to use: Internal or External.
            // AutoDiscover may not return an external protocol. First try external, then internal.
            // Either protocol may be returned without a configured URL.
            if ((isExternal &&
                response.TryGetSettingValue<string>(UserSettingName.ExternalEwsUrl, out uriString)) &&
                !string.IsNullOrEmpty(uriString))
            {
                return new Uri(uriString);
            }
            else if ((response.TryGetSettingValue<string>(UserSettingName.InternalEwsUrl, out uriString) ||
                     response.TryGetSettingValue<string>(UserSettingName.ExternalEwsUrl, out uriString)) &&
                     !string.IsNullOrEmpty(uriString))
            {
                return new Uri(uriString);
            }

            // If Autodiscover doesn't return an internal or external EWS URL, throw an exception.
            throw new AutodiscoverLocalException(Strings.AutodiscoverDidNotReturnEwsUrl);
        }

        #endregion

        #region ClientAccessTokens

        /// <summary>
        /// GetClientAccessToken
        /// </summary>
        /// <param name="idAndTypes">Id and Types</param>
        /// <returns>A ServiceResponseCollection providing token results for each of the specified id and types.</returns>
        public ServiceResponseCollection<GetClientAccessTokenResponse> GetClientAccessToken(IEnumerable<KeyValuePair<string, ClientAccessTokenType>> idAndTypes)
        {
            GetClientAccessTokenRequest request = new GetClientAccessTokenRequest(this, ServiceErrorHandling.ReturnErrors);
            List<ClientAccessTokenRequest> requestList = new List<ClientAccessTokenRequest>();
            foreach (KeyValuePair<string, ClientAccessTokenType> idAndType in idAndTypes)
            {
                ClientAccessTokenRequest clientAccessTokenRequest = new ClientAccessTokenRequest(idAndType.Key, idAndType.Value);
                requestList.Add(clientAccessTokenRequest);
            }

            return this.GetClientAccessToken(requestList.ToArray());
        }

        /// <summary>
        /// GetClientAccessToken
        /// </summary>
        /// <param name="tokenRequests">Token requests array</param>
        /// <returns>A ServiceResponseCollection providing token results for each of the specified id and types.</returns>
        public ServiceResponseCollection<GetClientAccessTokenResponse> GetClientAccessToken(ClientAccessTokenRequest[] tokenRequests)
        {
            GetClientAccessTokenRequest request = new GetClientAccessTokenRequest(this, ServiceErrorHandling.ReturnErrors);
            request.TokenRequests = tokenRequests;
            return request.Execute();
        }

        #endregion

        #region Client Extensibility

        /// <summary>
        /// Get the app manifests.
        /// </summary>
        /// <returns>Collection of manifests</returns>
        public Collection<XmlDocument> GetAppManifests()
        {
            GetAppManifestsRequest request = new GetAppManifestsRequest(this);
            return request.Execute().Manifests;
        }

        /// <summary>
        /// Get the app manifests.  Works with Exchange 2013 SP1 or later EWS.
        /// </summary>
        /// <param name="apiVersionSupported">The api version supported by the client.</param>
        /// <param name="schemaVersionSupported">The schema version supported by the client.</param>
        /// <returns>Collection of manifests</returns>
        public Collection<ClientApp> GetAppManifests(string apiVersionSupported, string schemaVersionSupported)
        {
            GetAppManifestsRequest request = new GetAppManifestsRequest(this);
            request.ApiVersionSupported = apiVersionSupported;
            request.SchemaVersionSupported = schemaVersionSupported;

            return request.Execute().Apps;
        }

        /// <summary>
        /// Install App. 
        /// </summary>
        /// <param name="manifestStream">The manifest's plain text XML stream. 
        /// Notice: Stream has state. If you want this function read from the expected position of the stream,
        /// please make sure set read position by manifestStream.Position = expectedPosition.
        /// Be aware read manifestStream.Lengh puts stream's Position at stream end.
        /// If you retrieve manifestStream.Lengh before call this function, nothing will be read.
        /// When this function succeeds, manifestStream is closed. This is by EWS design to 
        /// release resource in timely manner. </param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public void InstallApp(Stream manifestStream)
        {
            EwsUtilities.ValidateParam(manifestStream, "manifestStream");

            this.InternalInstallApp(manifestStream, marketplaceAssetId: null, marketplaceContentMarket: null, sendWelcomeEmail: false);
        }

        /// <summary>
        /// Install App. 
        /// </summary>
        /// <param name="manifestStream">The manifest's plain text XML stream. 
        /// Notice: Stream has state. If you want this function read from the expected position of the stream,
        /// please make sure set read position by manifestStream.Position = expectedPosition.
        /// Be aware read manifestStream.Lengh puts stream's Position at stream end.
        /// If you retrieve manifestStream.Lengh before call this function, nothing will be read.
        /// When this function succeeds, manifestStream is closed. This is by EWS design to 
        /// release resource in timely manner. </param>
        /// <param name="marketplaceAssetId">The asset id of the addin in marketplace</param>
        /// <param name="marketplaceContentMarket">The target market for content</param>
        /// <param name="sendWelcomeEmail">Whether to send welcome email for the addin</param>
        /// <returns>True if the app was not already installed. False if it was not installed. Null if it is not a user mailbox.</returns>
        /// <remarks>Exception will be thrown for errors. </remarks>
        internal bool? InternalInstallApp(Stream manifestStream, string marketplaceAssetId, string marketplaceContentMarket, bool sendWelcomeEmail)
        {
            EwsUtilities.ValidateParam(manifestStream, "manifestStream");

            InstallAppRequest request = new InstallAppRequest(this, manifestStream, marketplaceAssetId, marketplaceContentMarket, false);

            InstallAppResponse response = request.Execute();

            return response.WasFirstInstall;
        }

        /// <summary>
        /// Uninstall app. 
        /// </summary>
        /// <param name="id">App ID</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public void UninstallApp(string id)
        {
            EwsUtilities.ValidateParam(id, "id");

            UninstallAppRequest request = new UninstallAppRequest(this, id);

            request.Execute();
        }

        /// <summary>
        /// Disable App.
        /// </summary>
        /// <param name="id">App ID</param>
        /// <param name="disableReason">Disable reason</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public void DisableApp(string id, DisableReasonType disableReason)
        {
            EwsUtilities.ValidateParam(id, "id");
            EwsUtilities.ValidateParam(disableReason, "disableReason");

            DisableAppRequest request = new DisableAppRequest(this, id, disableReason);

            request.Execute();
        }

        /// <summary>
        /// Sets the consent state of an extension.
        /// </summary>
        /// <param name="id">Extension id.</param>
        /// <param name="state">Sets the consent state of an extension.</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public void RegisterConsent(string id, ConsentState state)
        {
            EwsUtilities.ValidateParam(id, "id");
            EwsUtilities.ValidateParam(state, "state");

            RegisterConsentRequest request = new RegisterConsentRequest(this, id, state);

            request.Execute();
        }

        /// <summary>
        /// Get App Marketplace Url.
        /// </summary>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public string GetAppMarketplaceUrl()
        {
            return GetAppMarketplaceUrl(null, null);
        }

        /// <summary>
        /// Get App Marketplace Url.  Works with Exchange 2013 SP1 or later EWS.
        /// </summary>
        /// <param name="apiVersionSupported">The api version supported by the client.</param>
        /// <param name="schemaVersionSupported">The schema version supported by the client.</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public string GetAppMarketplaceUrl(string apiVersionSupported, string schemaVersionSupported)
        {
            GetAppMarketplaceUrlRequest request = new GetAppMarketplaceUrlRequest(this);
            request.ApiVersionSupported = apiVersionSupported;
            request.SchemaVersionSupported = schemaVersionSupported;

            return request.Execute().AppMarketplaceUrl;
        }

        /// <summary>
        /// Get the client extension data. This method is used in server-to-server calls to retrieve ORG extensions for
        /// admin powershell/UMC access and user's powershell/UMC access as well as user's activation for OWA/Outlook.
        /// This is expected to never be used or called directly from user client.
        /// </summary>
        /// <param name="requestedExtensionIds">An array of requested extension IDs to return.</param>
        /// <param name="shouldReturnEnabledOnly">Whether enabled extension only should be returned, e.g. for user's
        /// OWA/Outlook activation scenario.</param>
        /// <param name="isUserScope">Whether it's called from admin or user scope</param>
        /// <param name="userId">Specifies optional (if called with user scope) user identity. This will allow to do proper
        /// filtering in cases where admin installs an extension for specific users only</param>
        /// <param name="userEnabledExtensionIds">Optional list of org extension IDs which user enabled. This is necessary for
        /// proper result filtering on the server end. E.g. if admin installed N extensions but didn't enable them, it does not
        /// make sense to return manifests for those which user never enabled either. Used only when asked
        /// for enabled extension only (activation scenario).</param>
        /// <param name="userDisabledExtensionIds">Optional list of org extension IDs which user disabled. This is necessary for
        /// proper result filtering on the server end. E.g. if admin installed N optional extensions and enabled them, it does
        /// not make sense to retrieve manifests for extensions which user disabled for him or herself. Used only when asked
        /// for enabled extension only (activation scenario).</param>
        /// <param name="isDebug">Optional flag to indicate whether it is debug mode. 
        /// If it is, org master table in arbitration mailbox will be returned for debugging purpose.</param>
        /// <returns>Collection of ClientExtension objects</returns>
        public GetClientExtensionResponse GetClientExtension(
            StringList requestedExtensionIds,
            bool shouldReturnEnabledOnly,
            bool isUserScope,
            string userId,
            StringList userEnabledExtensionIds,
            StringList userDisabledExtensionIds,
            bool isDebug)
        {
            GetClientExtensionRequest request = new GetClientExtensionRequest(
                this,
                requestedExtensionIds,
                shouldReturnEnabledOnly,
                isUserScope,
                userId,
                userEnabledExtensionIds,
                userDisabledExtensionIds,
                isDebug);

            return request.Execute();
        }

        /// <summary>
        /// Get the OME (i.e. Office Message Encryption) configuration data. This method is used in server-to-server calls to retrieve OME configuration
        /// </summary>
        /// <returns>OME Configuration response object</returns>
        public GetOMEConfigurationResponse GetOMEConfiguration()
        {
            GetOMEConfigurationRequest request = new GetOMEConfigurationRequest(this);

            return request.Execute();
        }

        /// <summary>
        /// Set the OME (i.e. Office Message Encryption) configuration data. This method is used in server-to-server calls to set encryption configuration
        /// </summary>
        /// <param name="xml">The xml</param>
        public void SetOMEConfiguration(string xml)
        {
            SetOMEConfigurationRequest request = new SetOMEConfigurationRequest(this, xml);

            request.Execute();
        }

        /// <summary>
        /// Set the client extension data. This method is used in server-to-server calls to install/uninstall/configure ORG
        /// extensions to support admin's management of ORG extensions via powershell/UMC.
        /// </summary>
        /// <param name="actions">List of actions to execute.</param>
        public void SetClientExtension(List<SetClientExtensionAction> actions)
        {
            SetClientExtensionRequest request = new SetClientExtensionRequest(this, actions);

            request.Execute();
        }

        #endregion

        #region Groups
        /// <summary>
        /// Gets the list of unified groups associated with the user
        /// </summary>
        /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
        /// <param name="userSmtpAddress">The smtp address of accessing user.</param>
        /// <returns>UserUnified groups.</returns>
        public Collection<UnifiedGroupsSet> GetUserUnifiedGroups(
                            IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets,
                            string userSmtpAddress)
        {
            EwsUtilities.ValidateParam(requestedUnifiedGroupsSets, "requestedUnifiedGroupsSets");
            EwsUtilities.ValidateParam(userSmtpAddress, "userSmtpAddress");

            return this.GetUserUnifiedGroupsInternal(requestedUnifiedGroupsSets, userSmtpAddress);
        }

        /// <summary>
        /// Gets the list of unified groups associated with the user
        /// </summary>
        /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
        /// <returns>UserUnified groups.</returns>
        public Collection<UnifiedGroupsSet> GetUserUnifiedGroups(IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets)
        {
            EwsUtilities.ValidateParam(requestedUnifiedGroupsSets, "requestedUnifiedGroupsSets");
            return this.GetUserUnifiedGroupsInternal(requestedUnifiedGroupsSets, null);
        }

        /// <summary>
        /// Gets the list of unified groups associated with the user
        /// </summary>
        /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
        /// <param name="userSmtpAddress">The smtp address of accessing user.</param>
        /// <returns>UserUnified groups.</returns>
        private Collection<UnifiedGroupsSet> GetUserUnifiedGroupsInternal(
                            IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets,
                            string userSmtpAddress)
        {
            GetUserUnifiedGroupsRequest request = new GetUserUnifiedGroupsRequest(this);

            if (!string.IsNullOrEmpty(userSmtpAddress))
            {
                request.UserSmtpAddress = userSmtpAddress;
            }

            if (requestedUnifiedGroupsSets != null)
            {
                request.RequestedUnifiedGroupsSets = requestedUnifiedGroupsSets;
            }

            return request.Execute().GroupsSets;
        }

        /// <summary>
        /// Gets the UnifiedGroupsUnseenCount for the group specfied 
        /// </summary>
        /// <param name="groupMailboxSmtpAddress">The smtpaddress of group for which unseendata is desired</param>
        /// <param name="lastVisitedTimeUtc">The LastVisitedTimeUtc of group for which unseendata is desired</param>
        /// <returns>UnifiedGroupsUnseenCount</returns>
        public int GetUnifiedGroupUnseenCount(string groupMailboxSmtpAddress, DateTime lastVisitedTimeUtc)
        {
            EwsUtilities.ValidateParam(groupMailboxSmtpAddress, "groupMailboxSmtpAddress");

            GetUnifiedGroupUnseenCountRequest request = new GetUnifiedGroupUnseenCountRequest(
                this, lastVisitedTimeUtc, UnifiedGroupIdentityType.SmtpAddress, groupMailboxSmtpAddress);

            request.AnchorMailbox = groupMailboxSmtpAddress;

            return request.Execute().UnseenCount;
        }

        /// <summary>
        /// Sets the LastVisitedTime for the group specfied 
        /// </summary>
        /// <param name="groupMailboxSmtpAddress">The smtpaddress of group for which unseendata is desired</param>
        /// <param name="lastVisitedTimeUtc">The LastVisitedTimeUtc of group for which unseendata is desired</param>
        public void SetUnifiedGroupLastVisitedTime(string groupMailboxSmtpAddress, DateTime lastVisitedTimeUtc)
        {
            EwsUtilities.ValidateParam(groupMailboxSmtpAddress, "groupMailboxSmtpAddress");

            SetUnifiedGroupLastVisitedTimeRequest request = new SetUnifiedGroupLastVisitedTimeRequest(this, lastVisitedTimeUtc, UnifiedGroupIdentityType.SmtpAddress, groupMailboxSmtpAddress);

            request.Execute();
        }

        #endregion

        #region Diagnostic Method -- Only used by test

        /// <summary>
        /// Executes the diagnostic method.
        /// </summary>
        /// <param name="verb">The verb.</param>
        /// <param name="parameter">The parameter.</param>
        /// <returns></returns>
        internal XmlDocument ExecuteDiagnosticMethod(string verb, XmlNode parameter)
        {
            ExecuteDiagnosticMethodRequest request = new ExecuteDiagnosticMethodRequest(this);
            request.Verb = verb;
            request.Parameter = parameter;

            return request.Execute()[0].ReturnValue;
        }
        #endregion

        #region Validation

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.Url == null)
            {
                throw new ServiceLocalException(Strings.ServiceUrlMustBeSet);
            }

            if (this.PrivilegedUserId != null && this.ImpersonatedUserId != null)
            {
                throw new ServiceLocalException(Strings.CannotSetBothImpersonatedAndPrivilegedUser);
            }

            // only one of PrivilegedUserId|ImpersonatedUserId|ManagementRoles can be set.
        }

        /// <summary>
        /// Validates a new-style version string.
        /// This validation is not as strict as server-side validation.
        /// </summary>
        /// <param name="version"> the version string </param>
        /// <remarks>
        /// The target version string has a required part and an optional part.
        /// The required part is two integers separated by a dot, major.minor
        /// The optional part is a minimum required version, minimum=major.minor
        /// Examples:
        ///   X-EWS-TargetVersion: 2.4
        ///   X-EWS_TargetVersion: 2.9; minimum=2.4
        /// </remarks>
        internal static void ValidateTargetVersion(string version)
        {
            const char ParameterSeparator = ';';
            const string LegacyVersionPrefix = "Exchange20";
            const char ParameterValueSeparator = '=';
            const string ParameterName = "minimum";

            if (String.IsNullOrEmpty(version))
            {
                throw new ArgumentException("Target version must not be empty.");
            }

            string[] parts = version.Trim().Split(ParameterSeparator);
            switch (parts.Length)
            {
                case 1:
                    // Validate the header value. We allow X.Y or Exchange20XX.
                    string part1 = parts[0].Trim();
                    if (parts[0].StartsWith(LegacyVersionPrefix))
                    {
                        // Close enough; misses corner cases like "Exchange2001". Server will do complete validation.
                    }
                    else if (ExchangeService.IsMajorMinor(part1))
                    {
                        // Also close enough; misses corner cases like ".5".
                    }
                    else
                    {
                        throw new ArgumentException("Target version must match X.Y or Exchange20XX.");
                    }

                    break;

                case 2:
                    // Validate the optional minimum version parameter, "minimum=X.Y"
                    string part2 = parts[1].Trim();
                    string[] minParts = part2.Split(ParameterValueSeparator);
                    if (minParts.Length == 2 &&
                        minParts[0].Trim().Equals(ParameterName, StringComparison.OrdinalIgnoreCase) &&
                        ExchangeService.IsMajorMinor(minParts[1].Trim()))
                    {
                        goto case 1;
                    }

                    throw new ArgumentException("Target version must match X.Y or Exchange20XX.");

                default:
                    throw new ArgumentException("Target version should have the form.");
            }
        }

        private static bool IsMajorMinor(string versionPart)
        {
            const char MajorMinorSeparator = '.';

            string[] parts = versionPart.Split(MajorMinorSeparator);
            if (parts.Length != 2)
            {
                return false;
            }

            foreach (string s in parts)
            {
                foreach (char c in s)
                {
                    if (!Char.IsDigit(c))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the latest supported version of EWS and scoped to the system's current time zone.
        /// </summary>
        public ExchangeService()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the latest supported version of EWS and scoped to the specified time zone.
        /// </summary>
        /// <param name="timeZone">The time zone to which the service is scoped.</param>
        public ExchangeService(TimeZoneInfo timeZone)
            : base(timeZone)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the specified version of EWS and scoped to the system's current time zone.
        /// </summary>
        /// <param name="requestedServerVersion">The version of EWS that the service targets.</param>
        public ExchangeService(ExchangeVersion requestedServerVersion)
            : base(requestedServerVersion)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the specified version of EWS and scoped to the specified time zone.
        /// </summary>
        /// <param name="requestedServerVersion">The version of EWS that the service targets.</param>
        /// <param name="timeZone">The time zone to which the service is scoped.</param>
        public ExchangeService(ExchangeVersion requestedServerVersion, TimeZoneInfo timeZone)
            : base(requestedServerVersion, timeZone)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the specified version of EWS and scoped to the system's current time zone.
        /// </summary>
        /// <param name="targetServerVersion">The version (new style) of EWS that the service targets.</param>
        /// <remarks>
        /// The target version string has a required part and an optional part.
        /// The required part is two integers separated by a dot, major.minor
        /// The optional part is a minimum required version, minimum=major.minor
        /// Examples:
        ///   X-EWS-TargetVersion: 2.4
        ///   X-EWS_TargetVersion: 2.9; minimum=2.4
        /// </remarks>
        internal ExchangeService(string targetServerVersion)
            : base(ExchangeVersion.Exchange2013)
        {
            ExchangeService.ValidateTargetVersion(targetServerVersion);
            this.TargetServerVersion = targetServerVersion;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the specified version of EWS and scoped to the specified time zone.
        /// </summary>
        /// <param name="targetServerVersion">The version (new style) of EWS that the service targets.</param>
        /// <param name="timeZone">The time zone to which the service is scoped.</param>
        /// <remarks>
        /// The new style version string has a required part and an optional part.
        /// The required part is two integers separated by a dot, major.minor
        /// The optional part is a minimum required version, minimum=major.minor
        /// Examples:
        ///   2.4
        ///   2.9; minimum=2.4
        /// </remarks>
        internal ExchangeService(string targetServerVersion, TimeZoneInfo timeZone)
            : base(ExchangeVersion.Exchange2013, timeZone)
        {
            ExchangeService.ValidateTargetVersion(targetServerVersion);
            this.TargetServerVersion = targetServerVersion;
        }

        #endregion

        #region Utilities
        /// <summary>
        /// Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
        /// based on the configuration of this service object.
        /// </summary>
        /// <param name="methodName">Name of the method.</param>
        /// <returns>
        /// An initialized instance of HttpWebRequest.
        /// </returns>
        internal IEwsHttpWebRequest PrepareHttpWebRequest(string methodName)
        {
            Uri endpoint = this.Url;
            this.RegisterCustomBasicAuthModule();

            endpoint = this.AdjustServiceUriFromCredentials(endpoint);

            IEwsHttpWebRequest request = this.PrepareHttpWebRequestForUrl(
                endpoint,
                this.AcceptGzipEncoding,
                true);

            if (!String.IsNullOrEmpty(this.TargetServerVersion))
            {
                request.Headers.Set(ExchangeService.TargetServerVersionHeaderName, this.TargetServerVersion);
            }

            return request;
        }

        /// <summary>
        /// Sets the type of the content.
        /// </summary>
        /// <param name="request">The request.</param>
        internal override void SetContentType(IEwsHttpWebRequest request)
        {
            request.ContentType = "text/xml; charset=utf-8";
            request.Accept = "text/xml";
        }

        /// <summary>
        /// Processes an HTTP error response.
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        /// <param name="webException">The web exception.</param>
        internal override void ProcessHttpErrorResponse(IEwsHttpWebResponse httpWebResponse, WebException webException)
        {
            this.InternalProcessHttpErrorResponse(
                httpWebResponse,
                webException,
                TraceFlags.EwsResponseHttpHeaders,
                TraceFlags.EwsResponse);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the URL of the Exchange Web Services. 
        /// </summary>
        public Uri Url
        {
            get { return this.url; }
            set { this.url = value; }
        }

        /// <summary>
        /// Gets or sets the Id of the user that EWS should impersonate. 
        /// </summary>
        public ImpersonatedUserId ImpersonatedUserId
        {
            get { return this.impersonatedUserId; }
            set { this.impersonatedUserId = value; }
        }

        /// <summary>
        /// Gets or sets the Id of the user that EWS should open his/her mailbox with privileged logon type. 
        /// </summary>
        internal PrivilegedUserId PrivilegedUserId
        {
            get { return this.privilegedUserId; }
            set { this.privilegedUserId = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public ManagementRoles ManagementRoles
        {
            get { return this.managementRoles; }
            set { this.managementRoles = value; }
        }

        /// <summary>
        /// Gets or sets the preferred culture for messages returned by the Exchange Web Services.
        /// </summary>
        public CultureInfo PreferredCulture
        {
            get { return this.preferredCulture; }
            set { this.preferredCulture = value; }
        }

        /// <summary>
        /// Gets or sets the DateTime precision for DateTime values returned from Exchange Web Services.
        /// </summary>
        public DateTimePrecision DateTimePrecision
        {
            get { return this.dateTimePrecision; }
            set { this.dateTimePrecision = value; }
        }

        /// <summary>
        /// Gets or sets a file attachment content handler.
        /// </summary>
        public IFileAttachmentContentHandler FileAttachmentContentHandler
        {
            get { return this.fileAttachmentContentHandler; }
            set { this.fileAttachmentContentHandler = value; }
        }

        /// <summary>
        /// Gets the time zone this service is scoped to.
        /// </summary>
        public new TimeZoneInfo TimeZone
        {
            get { return base.TimeZone; }
        }

        /// <summary>
        /// Provides access to the Unified Messaging functionalities.
        /// </summary>
        public UnifiedMessaging UnifiedMessaging
        {
            get
            {
                if (this.unifiedMessaging == null)
                {
                    this.unifiedMessaging = new UnifiedMessaging(this);
                }

                return this.unifiedMessaging;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the AutodiscoverUrl method should perform SCP (Service Connection Point) record lookup when determining
        /// the Autodiscover service URL.
        /// </summary>
        public bool EnableScpLookup
        {
            get { return this.enableScpLookup; }
            set { this.enableScpLookup = value; }
        }

        /// <summary>
        /// Exchange 2007 compatibility mode flag. (Off by default)
        /// </summary>
        private bool exchange2007CompatibilityMode;

        /// <summary>
        /// Gets or sets a value indicating whether Exchange2007 compatibility mode is enabled.
        /// </summary>
        /// <remarks>
        /// In order to support E12 servers, the Exchange2007CompatibilityMode property can be used 
        /// to indicate that we should use "Exchange2007" as the server version string rather than 
        /// Exchange2007_SP1.
        /// </remarks>
        internal bool Exchange2007CompatibilityMode
        {
            get { return this.exchange2007CompatibilityMode; }
            set { this.exchange2007CompatibilityMode = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether trace output is pretty printed.
        /// </summary>
        public bool TraceEnablePrettyPrinting
        {
            get { return this.traceEnablePrettyPrinting; }
            set { this.traceEnablePrettyPrinting = value; }
        }

        /// <summary>
        /// Gets or sets the target server version string (newer than Exchange2013).
        /// </summary>
        internal string TargetServerVersion
        {
            get
            {
                return this.targetServerVersion;
            }

            set
            {
                ExchangeService.ValidateTargetVersion(value);
                this.targetServerVersion = value;
            }
        }

        #endregion
    }
}