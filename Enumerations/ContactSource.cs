// ---------------------------------------------------------------------------
// <copyright file="ContactSource.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactSource enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the source of a contact or group.
    /// </summary>
    public enum ContactSource
    {
        /// <summary>
        /// The contact or group is stored in the Global Address List
        /// </summary>
        ActiveDirectory,

        /// <summary>
        /// The contact or group is stored in Exchange.
        /// </summary>
        Store
    }
}
