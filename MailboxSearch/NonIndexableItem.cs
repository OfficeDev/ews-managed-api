#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the NonIndexableItem class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Item index error
    /// </summary>
    public enum ItemIndexError
    {
        /// <summary>
        /// None
        /// </summary>
        None,

        /// <summary>
        /// Generic error
        /// </summary>
        GenericError,

        /// <summary>
        /// Timeout
        /// </summary>
        Timeout,

        /// <summary>
        /// Stale event
        /// </summary>
        StaleEvent,

        /// <summary>
        /// Mailbox offline
        /// </summary>
        MailboxOffline,

        /// <summary>
        /// Too many attachments to index
        /// </summary>
        AttachmentLimitReached,

        /// <summary>
        /// Data is truncated
        /// </summary>
        MarsWriterTruncation,
    }

    /// <summary>
    /// Represents non indexable item.
    /// </summary>
    public sealed class NonIndexableItem
    {
        /// <summary>
        /// Item Identity
        /// </summary>
        public ItemId ItemId { get; set; }

        /// <summary>
        /// Error code
        /// </summary>
        public ItemIndexError ErrorCode { get; set; }

        /// <summary>
        /// Error description
        /// </summary>
        public string ErrorDescription { get; set; }

        /// <summary>
        /// Is partially indexed
        /// </summary>
        public bool IsPartiallyIndexed { get; set; }

        /// <summary>
        /// Is permanent failure
        /// </summary>
        public bool IsPermanentFailure { get; set; }

        /// <summary>
        /// Attempt count
        /// </summary>
        public int AttemptCount { get; set; }

        /// <summary>
        /// Last attempt time
        /// </summary>
        public DateTime? LastAttemptTime { get; set; }

        /// <summary>
        /// Additional info
        /// </summary>
        public string AdditionalInfo { get; set; }

        /// <summary>
        /// Sort value
        /// </summary>
        public string SortValue { get; set; }

        /// <summary>
        /// Load from xml
        /// </summary>
        /// <param name="reader">The reader</param>
        /// <returns>Non indexable item object</returns>
        internal static NonIndexableItem LoadFromXml(EwsServiceXmlReader reader)
        {
            NonIndexableItem result = null;
            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.NonIndexableItemDetail))
            {
                ItemId itemId = null;
                ItemIndexError errorCode = ItemIndexError.None;
                string errorDescription = null;
                bool isPartiallyIndexed = false;
                bool isPermanentFailure = false;
                int attemptCount = 0;
                DateTime? lastAttemptTime = null;
                string additionalInfo = null;
                string sortValue = null;

                do
                {
                    reader.Read();
                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ItemId))
                    {
                        itemId = new ItemId();
                        itemId.ReadAttributesFromXml(reader);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ErrorDescription))
                    {
                        errorDescription = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ErrorDescription);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.IsPartiallyIndexed))
                    {
                        isPartiallyIndexed = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.IsPartiallyIndexed);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.IsPermanentFailure))
                    {
                        isPermanentFailure = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.IsPermanentFailure);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.AttemptCount))
                    {
                        attemptCount = reader.ReadElementValue<int>(XmlNamespace.Types, XmlElementNames.AttemptCount);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.LastAttemptTime))
                    {
                        lastAttemptTime = reader.ReadElementValue<DateTime>(XmlNamespace.Types, XmlElementNames.LastAttemptTime);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.AdditionalInfo))
                    {
                        additionalInfo = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.AdditionalInfo);
                    }
                    else if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SortValue))
                    {
                        sortValue = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SortValue);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.NonIndexableItemDetail));

                result = new NonIndexableItem
                {
                    ItemId = itemId,
                    ErrorCode = errorCode,
                    ErrorDescription = errorDescription,
                    IsPartiallyIndexed = isPartiallyIndexed,
                    IsPermanentFailure = isPermanentFailure,
                    AttemptCount = attemptCount,
                    LastAttemptTime = lastAttemptTime,
                    AdditionalInfo = additionalInfo,
                    SortValue = sortValue,
                };
            }

            return result;
        }
    }
}
