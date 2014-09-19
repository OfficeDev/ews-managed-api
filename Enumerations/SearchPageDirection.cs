// ---------------------------------------------------------------------------
// <copyright file="SearchPageDirection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchPageDirection enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the page direction for mailbox search.
    /// </summary>
    public enum SearchPageDirection
    {
        /// <summary>
        /// Navigate to next page.
        /// </summary>
        Next,

        /// <summary>
        /// Navigate to previous page.
        /// </summary>
        Previous
    }
}
