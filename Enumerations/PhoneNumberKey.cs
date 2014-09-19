// ---------------------------------------------------------------------------
// <copyright file="PhoneNumberKey.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhoneNumberKey enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines phone number entries for a contact.
    /// </summary>
    public enum PhoneNumberKey
    {
        /// <summary>
        /// The assistant's phone number.
        /// </summary>
        AssistantPhone,

        /// <summary>
        /// The business fax number.
        /// </summary>
        BusinessFax,

        /// <summary>
        /// The business phone number.
        /// </summary>
        BusinessPhone,

        /// <summary>
        /// The second business phone number.
        /// </summary>
        BusinessPhone2,

        /// <summary>
        /// The callback number.
        /// </summary>
        Callback,

        /// <summary>
        /// The car phone number.
        /// </summary>
        CarPhone,

        /// <summary>
        /// The company's main phone number.
        /// </summary>
        CompanyMainPhone,

        /// <summary>
        /// The home fax number.
        /// </summary>
        HomeFax,

        /// <summary>
        /// The home phone number.
        /// </summary>
        HomePhone,

        /// <summary>
        /// The second home phone number.
        /// </summary>
        HomePhone2,

        /// <summary>
        /// The ISDN number.
        /// </summary>
        Isdn,

        /// <summary>
        /// The mobile phone number.
        /// </summary>
        MobilePhone,

        /// <summary>
        /// An alternate fax number.
        /// </summary>
        OtherFax,

        /// <summary>
        /// An alternate phone number.
        /// </summary>
        OtherTelephone,

        /// <summary>
        /// The pager number.
        /// </summary>
        Pager,

        /// <summary>
        /// The primary phone number.
        /// </summary>
        PrimaryPhone,

        /// <summary>
        /// The radio phone number.
        /// </summary>
        RadioPhone,

        /// <summary>
        /// The Telex number.
        /// </summary>
        Telex,

        /// <summary>
        /// The TTY/TTD phone number.
        /// </summary>
        TtyTddPhone
    }
}
