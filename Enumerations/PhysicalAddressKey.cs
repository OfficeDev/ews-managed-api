// ---------------------------------------------------------------------------
// <copyright file="PhysicalAddressKey.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhysicalAddressKey enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines physical address entries for a contact.
    /// </summary>
    public enum PhysicalAddressKey
    {
        /// <summary>
        /// The business address.
        /// </summary>
        Business,

        /// <summary>
        /// The home address.
        /// </summary>
        Home,

        /// <summary>
        /// An alternate address.
        /// </summary>
        Other
    }
}
