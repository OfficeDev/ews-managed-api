// ---------------------------------------------------------------------------
// <copyright file="SearchMailboxesParameters.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchMailboxesParameters class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents search mailbox parameters.
    /// </summary>
    public sealed class SearchMailboxesParameters
    {
        /// <summary>
        /// Search queries
        /// </summary>
        public MailboxQuery[] SearchQueries { get; set; }

        /// <summary>
        /// Result type
        /// </summary>
        public SearchResultType ResultType { get; set; }

        /// <summary>
        /// Sort by property
        /// </summary>
        public string SortBy { get; set; }

        /// <summary>
        /// Sort direction
        /// </summary>
        public SortDirection SortOrder { get; set; }

        /// <summary>
        /// Perform deduplication
        /// </summary>
        public bool PerformDeduplication { get; set; }

        /// <summary>
        /// Page size
        /// </summary>
        public int PageSize { get; set; }

        /// <summary>
        /// Search page direction
        /// </summary>
        public SearchPageDirection PageDirection { get; set; }

        /// <summary>
        /// Page item reference
        /// </summary>
        public string PageItemReference { get; set; }

        /// <summary>
        /// Preview item response shape
        /// </summary>
        public PreviewItemResponseShape PreviewItemResponseShape { get; set; }

        /// <summary>
        /// Query language
        /// </summary>
        public string Language { get; set; }
    }
}