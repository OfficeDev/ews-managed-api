// ---------------------------------------------------------------------------
// <copyright file="TeamMailboxLifecycleState.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TeamMailbox lifecycle state enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// TeamMailbox lifecycle state
    /// </summary>
    public enum TeamMailboxLifecycleState
    {
        /// <summary>
        /// Active
        /// </summary>
        [EwsEnum("Active")]
        Active,

        /// <summary>
        /// Closed
        /// </summary>
        [EwsEnum("Closed")]
        Closed,

        /// <summary>
        /// Unlinked
        /// </summary>
        [EwsEnum("Unlinked")]
        Unlinked,

        /// <summary>
        /// PendingDelete
        /// </summary>
        [EwsEnum("PendingDelete")]
        PendingDelete,
    }
}