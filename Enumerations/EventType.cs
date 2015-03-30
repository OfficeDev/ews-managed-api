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