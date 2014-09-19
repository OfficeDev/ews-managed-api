// ---------------------------------------------------------------------------
// <copyright file="SendInvitationsOrCancellationsMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SendInvitationsOrCancellationsMode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines if/how meeting invitations or cancellations should be sent to attendees when an appointment is updated.
    /// </summary>
    public enum SendInvitationsOrCancellationsMode
    {
        /// <summary>
        /// No meeting invitation/cancellation is sent.
        /// </summary>
        SendToNone,

        /// <summary>
        /// Meeting invitations/cancellations are sent to all attendees.
        /// </summary>
        SendOnlyToAll,

        /// <summary>
        /// Meeting invitations/cancellations are sent only to attendees that have been added or modified.
        /// </summary>
        SendOnlyToChanged,

        /// <summary>
        /// Meeting invitations/cancellations are sent to all attendees and a copy is saved in the organizer's Sent Items folder.
        /// </summary>
        SendToAllAndSaveCopy,

        /// <summary>
        /// Meeting invitations/cancellations are sent only to attendees that have been added or modified and a copy is saved in the organizer's Sent Items folder.
        /// </summary>
        SendToChangedAndSaveCopy
    }
}
