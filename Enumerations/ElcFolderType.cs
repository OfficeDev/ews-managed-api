// ---------------------------------------------------------------------------
// <copyright file="ElcFolderType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ElcFolderType enumeration.</summary>
//-----------------------------------------------------------------------

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the folder type of a retention policy tag.
    /// </summary>
    public enum ElcFolderType
    {
        /// <summary>
        /// Calendar folder.
        /// </summary>
        Calendar = 1,

        /// <summary>
        /// Contacts folder.
        /// </summary>
        Contacts = 2,

        /// <summary>
        /// Deleted Items.
        /// </summary>
        DeletedItems = 3,

        /// <summary>
        /// Drafts folder.
        /// </summary>
        Drafts = 4,

        /// <summary>
        /// Inbox.
        /// </summary>
        Inbox = 5,

        /// <summary>
        /// Junk mail.
        /// </summary>
        JunkEmail = 6,

        /// <summary>
        /// Journal.
        /// </summary>
        Journal = 7,

        /// <summary>
        /// Notes.
        /// </summary>
        Notes = 8,

        /// <summary>
        /// Outbox.
        /// </summary>
        Outbox = 9,

        /// <summary>
        /// Sent Items.
        /// </summary>
        SentItems = 10,

        /// <summary>
        /// Tasks folder.
        /// </summary>
        Tasks = 11,

        /// <summary>
        /// Policy applies to all folders that do not have a policy.
        /// </summary>
        All = 12,

        /// <summary>
        /// Policy is for an organizational policy.
        /// </summary>
        ManagedCustomFolder = 13,

        /// <summary>
        /// Policy is for the RSS Subscription (default) folder.
        /// </summary>
        RssSubscriptions = 14,

        /// <summary>
        /// Policy is for the Sync Issues (default) folder.
        /// </summary>
        SyncIssues = 15,

        /// <summary>
        /// Policy is for the Conversation History (default) folder.
        /// This folder is used by the Office Communicator to archive IM conversations.
        /// </summary>
        ConversationHistory = 16,

        /// <summary>
        /// Policy is for the personal folders.
        /// </summary>
        Personal = 17,

        /// <summary>
        /// Policy is for Dumpster 2.0.
        /// </summary>
        RecoverableItems = 18,

        /// <summary>
        /// Non IPM Subtree root.
        /// </summary>
        NonIpmRoot = 19,
    }
}
