// ---------------------------------------------------------------------------
// <copyright file="NonIndexableItemParameters.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NonIndexableItemParameters class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents non indexable item parameters base class
    /// </summary>
    public abstract class NonIndexableItemParameters
    {
        /// <summary>
        /// List of mailboxes (in legacy DN format)
        /// </summary>
        public string[] Mailboxes { get; set; }

        /// <summary>
        /// Search archive only
        /// </summary>
        public bool SearchArchiveOnly { get; set; }
    }

    /// <summary>
    /// Represents get non indexable item statistics parameters.
    /// </summary>
    public sealed class GetNonIndexableItemStatisticsParameters : NonIndexableItemParameters
    {
    }

    /// <summary>
    /// Represents get non indexable item details parameters.
    /// </summary>
    public sealed class GetNonIndexableItemDetailsParameters : NonIndexableItemParameters
    {
        /// <summary>
        /// Page size
        /// </summary>
        public int? PageSize { get; set; }

        /// <summary>
        /// Page item reference
        /// </summary>
        public string PageItemReference { get; set; }

        /// <summary>
        /// Search page direction
        /// </summary>
        public SearchPageDirection? PageDirection { get; set; }
    }
}