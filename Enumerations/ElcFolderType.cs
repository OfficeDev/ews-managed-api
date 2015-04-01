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