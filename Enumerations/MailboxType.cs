// ---------------------------------------------------------------------------
// <copyright file="MailboxType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MailboxType enumeration.</summary>
//-----------------------------------------------------------------------

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
    }
}
