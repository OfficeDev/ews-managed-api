// ---------------------------------------------------------------------------
// <copyright file="SearchResultType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchResultType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of search result.
    /// </summary>
    public enum SearchResultType
    {
        /// <summary>
        /// Keyword statistics only.
        /// </summary>
        StatisticsOnly,

        /// <summary>
        /// Preview only.
        /// </summary>
        PreviewOnly,
    }
}
