// ---------------------------------------------------------------------------
// <copyright file="MailboxSearchScopeType.cs" company="Microsoft">
//     Copyright (c) Microsoft. All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// <summary>
//      MailboxSearchScopeType.cs
// </summary>
// ---------------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Enum MailboxSearchScopeType
    /// </summary>
    internal enum MailboxSearchScopeType
    {
        /// <summary>
        /// The legacy exchange DN
        /// </summary>
        LegacyExchangeDN = 0,

        /// <summary>
        /// The public folder
        /// </summary>
        PublicFolder = 1,

        /// <summary>
        /// The recipient
        /// </summary>
        Recipient = 2,

        /// <summary>
        /// The mailbox GUID
        /// </summary>
        MailboxGuid = 3,

        /// <summary>
        /// All public folders
        /// </summary>
        AllPublicFolders = 4,

        /// <summary>
        /// All mailboxes
        /// </summary>
        AllMailboxes = 5,

        /// <summary>
        /// The saved search id
        /// </summary>
        SavedSearchId = 6,

        /// <summary>
        /// The auto detect
        /// </summary>
        AutoDetect = 7,
    }
}
