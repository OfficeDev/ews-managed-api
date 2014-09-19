// ---------------------------------------------------------------------------
// <copyright file="ConversationSortOrder.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationSortOrder enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the order in which conversation nodes should be returned by GetConversationItems.
    /// </summary>
    public enum ConversationSortOrder
    {
        /// <summary>
        /// Tree order, ascending
        /// </summary>
        TreeOrderAscending,

        /// <summary>
        /// Tree order, descending.
        /// </summary>
        TreeOrderDescending,

        /// <summary>
        /// Chronological order, ascending.
        /// </summary>
        DateOrderAscending,

        /// <summary>
        /// Chronological order, descending.
        /// </summary>
        DateOrderDescending,
    }
}
