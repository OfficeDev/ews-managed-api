// ---------------------------------------------------------------------------
// <copyright file="EmailAddressKey.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailAddressKey enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines e-mail address entries for a contact.
    /// </summary>
    public enum EmailAddressKey
    {
        /// <summary>
        /// The first e-mail address.
        /// </summary>
        EmailAddress1,

        /// <summary>
        /// The second e-mail address.
        /// </summary>
        EmailAddress2,

        /// <summary>
        /// The third e-mail address.
        /// </summary>
        EmailAddress3
    }
}
