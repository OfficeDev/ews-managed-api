// ---------------------------------------------------------------------------
// <copyright file="ViewFilter.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ViewFilter enumeration.</summary>
//-----------------------------------------------------------------------

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the view filter for queries.
    /// </summary>
    public enum ViewFilter
    {
        /// <summary>
        /// Show all item (no filter)
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        All = 0,

        /// <summary>
        /// Item has flag
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        Flagged = 1,

        /// <summary>
        /// Item has attachment
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        HasAttachment = 2,

        /// <summary>
        /// Item is to or cc me
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        ToOrCcMe = 3,

        /// <summary>
        /// Item is unread
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        Unread = 4,

        /// <summary>
        /// Active task items
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        TaskActive = 5,

        /// <summary>
        /// Overdue task items
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        TaskOverdue = 6,

        /// <summary>
        /// Completed task items
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        TaskCompleted = 7,

        /// <summary>
        /// Suggestions (aka Predicted Actions) from the Inference engine
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        Suggestions = 8,

        /// <summary>
        /// Respond suggestions
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        SuggestionsRespond = 9,

        /// <summary>
        /// Delete suggestions
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2013)]
        SuggestionsDelete = 10,
    }
}