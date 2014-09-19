// ---------------------------------------------------------------------------
// <copyright file="PhysicalAddressIndex.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhysicalAddressType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines a physical address index.
    /// </summary>
    public enum PhysicalAddressIndex
    {
        /// <summary>
        /// None.
        /// </summary>
        None,

        /// <summary>
        /// The business address.
        /// </summary>
        Business,

        /// <summary>
        /// The home address.
        /// </summary>
        Home,

        /// <summary>
        /// The alternate address.
        /// </summary>
        Other
    }
}