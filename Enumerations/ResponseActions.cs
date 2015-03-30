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