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
    /// Represents failed mailbox to be searched
    /// </summary>
    public sealed class FailedSearchMailbox
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="mailbox">Mailbox identifier</param>
        /// <param name="errorCode">Error code</param>
        /// <param name="errorMessage">Error message</param>
        public FailedSearchMailbox(string mailbox, int errorCode, string errorMessage) :
            this(mailbox, errorCode, errorMessage, false)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="mailbox">Mailbox identifier</param>
        /// <param name="errorCode">Error code</param>
        /// <param name="errorMessage">Error message</param>
        /// <param name="isArchive">True if it is mailbox archive</param>
        public FailedSearchMailbox(string mailbox, int errorCode, string errorMessage, bool isArchive)
        {
            Mailbox = mailbox;
            ErrorCode = errorCode;
            ErrorMessage = errorMessage;
            IsArchive = isArchive;
        }

        /// <summary>
        /// Mailbox identifier
        /// </summary>
        public string Mailbox { get; set; }

        /// <summary>
        /// Error code
        /// </summary>
        public int ErrorCode { get; set; }

        /// <summary>
        /// Error message
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Whether it is archive mailbox or not
        /// </summary>
        public bool IsArchive { get; set; }

        /// <summary>
        /// Load failed mailboxes xml
        /// </summary>
        /// <param name="rootXmlNamespace">Root xml namespace</param>
        /// <param name="reader">The reader</param>
        /// <returns>Array of failed mailboxes</returns>
        internal static FailedSearchMailbox[] LoadFailedMailboxesXml(XmlNamespace rootXmlNamespace, EwsServiceXmlReader reader)
        {
            List<FailedSearchMailbox> failedMailboxes = new List<FailedSearchMailbox>();

            reader.EnsureCurrentNodeIsStartElement(rootXmlNamespace, XmlElementNames.FailedMailboxes);
            do
            {
                reader.Read();
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.FailedMailbox))
                {
                    string mailbox = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Mailbox);
                    int errorCode = 0;
                    int.TryParse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ErrorCode), out errorCode);
                    string errorMessage = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.ErrorMessage);
                    bool isArchive = false;
                    bool.TryParse(reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.IsArchive), out isArchive);

                    failedMailboxes.Add(new FailedSearchMailbox(mailbox, errorCode, errorMessage, isArchive));
                }
            }
            while (!reader.IsEndElement(rootXmlNamespace, XmlElementNames.FailedMailboxes));

            return failedMailboxes.Count == 0 ? null : failedMailboxes.ToArray();
        }
    }
}