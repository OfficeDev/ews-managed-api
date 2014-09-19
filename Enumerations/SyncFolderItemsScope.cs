// ---------------------------------------------------------------------------
// <copyright file="SyncFolderItemsScope.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SyncFolderItemsScope enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Determines items to be included in a SyncFolderItems response.
    /// </summary>
    public enum SyncFolderItemsScope
    {
        /// <summary>
        /// Include only normal items in the response.
        /// </summary>
        NormalItems,

        /// <summary>
        /// Include normal and associated items in the response.
        /// </summary>
        NormalAndAssociatedItems
    }
}
