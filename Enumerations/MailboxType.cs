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

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of an EmailAddress object.
    /// </summary>
    public enum MailboxType
    {
        /// <summary>
        /// Unknown mailbox type (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        Unknown,

        /// <summary>
        /// The EmailAddress represents a one-off contact (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        OneOff,

        /// <summary>
        /// The EmailAddress represents a mailbox.
        /// </summary>
        Mailbox,

        /// <summary>
        /// The EmailAddress represents a public folder.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2007_SP1)]
        PublicFolder,

        /// <summary>
        /// The EmailAddress represents a Public Group.
        /// </summary>
        [EwsEnumAttribute("PublicDL")]
        PublicGroup,

        /// <summary>
        /// The EmailAddress represents a Contact Group.
        /// </summary>
        [EwsEnumAttribute("PrivateDL")]
        ContactGroup,

        /// <summary>
        /// The EmailAddress represents a store contact or AD mail contact.
        /// </summary>
        Contact,

        /// <summary>
        /// The EmailAddress represents a GroupMailbox
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2016)]
        GroupMailbox,
    }
}