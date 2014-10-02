#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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