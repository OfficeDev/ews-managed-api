// ---------------------------------------------------------------------------
// <copyright file="ItemTraversal.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemTraversal enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the scope of FindItems operations.
    /// </summary>
    public enum ItemTraversal
    {
        /// <summary>
        /// All non deleted items in the specified folder are retrieved.
        /// </summary>
        Shallow,

        /// <summary>
        /// Only soft-deleted items are retrieved.
        /// </summary>
        SoftDeleted,

        /// <summary>
        /// Only associated items are retrieved (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        Associated
    }
}
