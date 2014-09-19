// ---------------------------------------------------------------------------
// <copyright file="SendCancellationsMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SendCancellationsMode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines how meeting cancellations should be sent to attendees when an appointment is deleted.
    /// </summary>
    public enum SendCancellationsMode
    {
        /// <summary>
        /// No meeting cancellation is sent.
        /// </summary>
        SendToNone,

        /// <summary>
        /// Meeting cancellations are sent to all attendees.
        /// </summary>
        SendOnlyToAll,

        /// <summary>
        /// Meeting cancellations are sent to all attendees and a copy of the cancellation message is saved in the organizer's Sent Items folder.
        /// </summary>
        SendToAllAndSaveCopy,
    }
}
