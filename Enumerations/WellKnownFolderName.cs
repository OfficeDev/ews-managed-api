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

    // The values in this enumeration must match the values of the
    // DistinguishedFolderIdNameType type in the schema.

    /// <summary>
    /// Defines well known folder names.
    /// </summary>
    public enum WellKnownFolderName
    {
        /// <summary>
        /// The Calendar folder.
        /// </summary>
        [EwsEnum("calendar")]
        Calendar,

        /// <summary>
        /// The Contacts folder.
        /// </summary>
        [EwsEnum("contacts")]
        Contacts,

        /// <summary>
        /// The Deleted Items folder
        /// </summary>
        [EwsEnum("deleteditems")]
        DeletedItems,

        /// <summary>
        /// The Drafts folder.
        /// </summary>
        [EwsEnum("drafts")]
        Drafts,

        /// <summary>
        /// The Inbox folder.
        /// </summary>
        [EwsEnum("inbox")]
        Inbox,

        /// <summary>
        /// The Journal folder.
        /// </summary>
        [EwsEnum("journal")]
        Journal,

        /// <summary>
        /// The Notes folder.
        /// </summary>
        [EwsEnum("notes")]
        Notes,

        /// <summary>
        /// The Outbox folder.
        /// </summary>
        [EwsEnum("outbox")]
        Outbox,

        /// <summary>
        /// The Sent Items folder.
        /// </summary>
        [EwsEnum("sentitems")]
        SentItems,

        /// <summary>
        /// The Tasks folder.
        /// </summary>
        [EwsEnum("tasks")]
        Tasks,

        /// <summary>
        /// The message folder root.
        /// </summary>
        [EwsEnum("msgfolderroot")]
        MsgFolderRoot,

        /// <summary>
        /// The root of the Public Folders hierarchy.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2007_SP1)]
        [EwsEnum("publicfoldersroot")]
        PublicFoldersRoot,

        /// <summary>
        /// The root of the mailbox.
        /// </summary>
        [EwsEnum("root")]
        Root,

        /// <summary>
        /// The Junk E-mail folder.
        /// </summary>
        [EwsEnum("junkemail")]
        JunkEmail,

        /// <summary>
        /// The Search Folders folder, also known as the Finder folder.
        /// </summary>
        [EwsEnum("searchfolders")]
        SearchFolders,

        /// <summary>
        /// The Voicemail folder.
        /// </summary>
        [EwsEnum("voicemail")]
        VoiceMail,

        /// <summary>
        /// The Dumpster 2.0 root folder.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("recoverableitemsroot")]
        RecoverableItemsRoot,

        /// <summary>
        /// The Dumpster 2.0 soft deletions folder.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("recoverableitemsdeletions")]
        RecoverableItemsDeletions,

        /// <summary>
        /// The Dumpster 2.0 versions folder.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("recoverableitemsversions")]
        RecoverableItemsVersions,

        /// <summary>
        /// The Dumpster 2.0 hard deletions folder.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("recoverableitemspurges")]
        RecoverableItemsPurges,
        
        /// <summary>
        /// The Dumpster 2.0 discovery hold folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013_SP1)]
        [EwsEnum("recoverableitemsdiscoveryholds")]
        RecoverableItemsDiscoveryHolds,

        /// <summary>
        /// The root of the archive mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archiveroot")]
        ArchiveRoot,

		/// <summary>
		/// The root of the archive mailbox.
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013_SP1)]
		[EwsEnum("archiveinbox")]
		ArchiveInbox,

		/// <summary>
		/// The message folder root in the archive mailbox.
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archivemsgfolderroot")]
        ArchiveMsgFolderRoot,

        /// <summary>
        /// The Deleted Items folder in the archive mailbox
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archivedeleteditems")]
        ArchiveDeletedItems,

        /// <summary>
        /// The Dumpster 2.0 root folder in the archive mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archiverecoverableitemsroot")]
        ArchiveRecoverableItemsRoot,

        /// <summary>
        /// The Dumpster 2.0 soft deletions folder in the archive mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archiverecoverableitemsdeletions")]
        ArchiveRecoverableItemsDeletions,

        /// <summary>
        /// The Dumpster 2.0 versions folder in the archive mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archiverecoverableitemsversions")]
        ArchiveRecoverableItemsVersions,

        /// <summary>
        /// The Dumpster 2.0 hard deletions folder in the archive mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("archiverecoverableitemspurges")]
        ArchiveRecoverableItemsPurges,

        /// <summary>
        /// The Dumpster 2.0 discovery hold folder in the archive mailbox.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013_SP1)]
        [EwsEnum("archiverecoverableitemsdiscoveryholds")]
        ArchiveRecoverableItemsDiscoveryHolds,

        /// <summary>
        /// The Sync Issues folder.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("syncissues")]
        SyncIssues,

        /// <summary>
        /// The Conflicts folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("conflicts")]
        Conflicts,

        /// <summary>
        /// The Local failures folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("localfailures")]
        LocalFailures,

        /// <summary>
        /// The Server failures folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("serverfailures")]
        ServerFailures,

        /// <summary>
        /// The Recipient Cache folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("recipientcache")]
        RecipientCache,

        /// <summary>
        /// The Quick Contacts folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("quickcontacts")]
        QuickContacts,

        /// <summary>
        /// Conversation history folder
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("conversationhistory")]
        ConversationHistory,

		/// <summary>
		/// AdminAuditLogs folder
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013)]
		[EwsEnum("adminauditlogs")]
		AdminAuditLogs,

		/// <summary>
		/// ToDo search folder
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013)]
        [EwsEnum("todosearch")]
        ToDoSearch,

		/// <summary>
		/// MyContacts folder
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013)]
		[EwsEnum("mycontacts")]
		MyContacts,

		/// <summary>
		/// Directory (GAL)
		/// It is not a mailbox folder. It only indicates any GAL operation.
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013_SP1)]
        [EwsEnum("directory")]
        Directory,

		/// <summary>
		/// IMContactList folder
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013)]
		[EwsEnum("imcontactlist")]
		IMContactList,

		/// <summary>
		/// PeopleConnect folder
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013)]
		[EwsEnum("peopleconnect")]
		PeopleConnect,

		/// <summary>
		/// Favorites folder
		/// </summary>
		[RequiredServerVersion(ExchangeVersion.Exchange2013)]
		[EwsEnum("favorites")]
		Favorites,

		//// Note when you adding new folder id here, please update sources\test\Services\src\ComponentTests\GlobalVersioningControl.cs
		//// IsExchange2013Folder method accordingly.
	}
}