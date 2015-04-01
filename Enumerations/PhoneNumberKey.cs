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