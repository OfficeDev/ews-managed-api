// ---------------------------------------------------------------------------
// <copyright file="ResponseActions.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseActions enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the response actions that can be taken on an item.
    /// </summary>
    [Flags]
    public enum ResponseActions
    {
        /// <summary>
        /// No action can be taken.
        /// </summary>
        None = 0,

        /// <summary>
        /// The item can be accepted.
        /// </summary>
        Accept = 1,

        /// <summary>
        /// The item can be tentatively accepted.
        /// </summary>
        TentativelyAccept = 2,

        /// <summary>
        /// The item can be declined.
        /// </summary>
        Decline = 4,

        /// <summary>
        /// The item can be replied to.
        /// </summary>
        Reply = 8,

        /// <summary>
        /// The item can be replied to.
        /// </summary>
        ReplyAll = 16,

        /// <summary>
        /// The item can be forwarded.
        /// </summary>
        Forward = 32,

        /// <summary>
        /// The item can be cancelled.
        /// </summary>
        Cancel = 64,

        /// <summary>
        /// The item can be removed from the calendar.
        /// </summary>
        RemoveFromCalendar = 128,

        /// <summary>
        /// The item's read receipt can be suppressed.
        /// </summary>
        SuppressReadReceipt = 256,

        /// <summary>
        /// A reply to the item can be posted.
        /// </summary>
        PostReply = 512
    }
}