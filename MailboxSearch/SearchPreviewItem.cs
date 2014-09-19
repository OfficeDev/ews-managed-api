// ---------------------------------------------------------------------------
// <copyright file="SearchPreviewItem.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchPreviewItem class.</summary>
//-----------------------------------------------------------------------

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
