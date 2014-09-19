// ---------------------------------------------------------------------------
// <copyright file="EventType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EventType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the types of event that can occur in a folder.
    /// </summary>
    public enum EventType
    {
        /// <summary>
        /// This event is sent to a client application by push notifications to indicate that
        /// the subscription is still alive.
        /// </summary>
        [EwsEnum("StatusEvent")]
        Status,

        /// <summary>
        /// This event indicates that a new e-mail message was received.
        /// </summary>
        [EwsEnum("NewMailEvent")]
        NewMail,

        /// <summary>
        /// This event indicates that an item or folder has been deleted.
        /// </summary>
        [EwsEnum("DeletedEvent")]
        Deleted,

        /// <summary>
        /// This event indicates that an item or folder has been modified.
        /// </summary>
        [EwsEnum("ModifiedEvent")]
        Modified,

        /// <summary>
        /// This event indicates that an item or folder has been moved to another folder.
        /// </summary>
        [EwsEnum("MovedEvent")]
        Moved,

        /// <summary>
        /// This event indicates that an item or folder has been copied to another folder.
        /// </summary>
        [EwsEnum("CopiedEvent")]
        Copied,

        /// <summary>
        /// This event indicates that a new item or folder has been created.
        /// </summary>
        [EwsEnum("CreatedEvent")]
        Created,

        /// <summary>
        /// This event indicates that free/busy has changed. This is only supported in 2010 SP1 or later
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010_SP1)]
        [EwsEnum("FreeBusyChangedEvent")]
        FreeBusyChanged
    }
}
