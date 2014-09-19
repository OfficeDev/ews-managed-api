// ---------------------------------------------------------------------------
// <copyright file="SendInvitationsMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SendInvitationsMode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines if/how meeting invitations are sent.
    /// </summary>
    public enum SendInvitationsMode
    {
        /// <summary>
        /// No meeting invitation is sent.
        /// </summary>
        SendToNone,

        /// <summary>
        /// Meeting invitations are sent to all attendees.
        /// </summary>
        SendOnlyToAll,

        /// <summary>
        /// Meeting invitations are sent to all attendees and a copy of the invitation message is saved.
        /// </summary>
        SendToAllAndSaveCopy
    }
}
