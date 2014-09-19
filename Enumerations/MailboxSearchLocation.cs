// ---------------------------------------------------------------------------
// <copyright file="MailboxSearchLocation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MailboxSearchLocation enumeration.</summary>
//-----------------------------------------------------------------------

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the location for mailbox search.
    /// </summary>
    public enum MailboxSearchLocation
    {
        /// <summary>
        /// Primary only (Exchange 2013 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        PrimaryOnly,

        /// <summary>
        /// Archive only (Exchange 2013 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        ArchiveOnly,

        /// <summary>
        /// Both Primary and Archive (Exchange 2013 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        All,
    }
}