// ---------------------------------------------------------------------------
// <copyright file="KeywordStatisticsSearchResult.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the KeywordStatisticsSearchResult class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the keyword statistics result.
    /// </summary>
    public sealed class KeywordStatisticsSearchResult
    {
        /// <summary>
        /// Keyword string
        /// </summary>
        public string Keyword { get; set; }

        /// <summary>
        /// Number of item hits
        /// </summary>
        public int ItemHits { get; set; }

        /// <summary>
        /// Total size
        /// </summary>
        [CLSCompliant(false)]
        public ulong Size { get; set; }
    }
}
