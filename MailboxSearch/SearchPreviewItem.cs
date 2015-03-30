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
    /// Represents search preview item.
    /// </summary>
    public sealed class SearchPreviewItem
    {
        /// <summary>
        /// Item id
        /// </summary>
        public ItemId Id { get; set; }

        /// <summary>
        /// Mailbox
        /// </summary>
        public PreviewItemMailbox Mailbox { get; set; }

        /// <summary>
        /// Parent item id
        /// </summary>
        public ItemId ParentId { get; set; }
        
        /// <summary>
        /// Item class
        /// </summary>
        public string ItemClass { get; set; }

        /// <summary>
        /// Unique hash
        /// </summary>
        public string UniqueHash { get; set; }

        /// <summary>
        /// Sort value
        /// </summary>
        public string SortValue { get; set; }

        /// <summary>
        /// OWA Link
        /// </summary>
        public string OwaLink { get; set; }

        /// <summary>
        /// Sender
        /// </summary>
        public string Sender { get; set; }

        /// <summary>
        /// To recipients
        /// </summary>
        public string[] ToRecipients { get; set; }

        /// <summary>
        /// Cc recipients
        /// </summary>
        public string[] CcRecipients { get; set; }

        /// <summary>
        /// Bcc recipients
        /// </summary>
        public string[] BccRecipients { get; set; }

        /// <summary>
        /// Created time
        /// </summary>
        public DateTime CreatedTime { get; set; }

        /// <summary>
        /// Received time
        /// </summary>
        public DateTime ReceivedTime { get; set; }

        /// <summary>
        /// Sent time
        /// </summary>
        public DateTime SentTime { get; set; }

        /// <summary>
        /// Subject
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Item size
        /// </summary>
        [CLSCompliant(false)]
        public ulong Size { get; set; }

        /// <summary>
        /// Preview
        /// </summary>
        public string Preview { get; set; }

        /// <summary>
        /// Importance
        /// </summary>
        public Importance Importance { get; set; }

        /// <summary>
        /// Read
        /// </summary>
        public bool Read { get; set; }

        /// <summary>
        /// Has attachments
        /// </summary>
        public bool HasAttachment { get; set; }

        /// <summary>
        /// Extended properties
        /// </summary>
        public ExtendedPropertyCollection ExtendedProperties { get; set; }
    }
}