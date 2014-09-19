// ---------------------------------------------------------------------------
// <copyright file="FindConversationResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindConversationResults class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents the results of an conversation search operation.
    /// </summary>
    public sealed class FindConversationResults
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FindConversationResults"/> class.
        /// </summary>
        internal FindConversationResults()
        {
            this.Conversations = new Collection<Conversation>();
            this.HighlightTerms = new Collection<HighlightTerm>();
            this.TotalCount = null;
       }

        /// <summary>
        /// Gets a collection containing the conversations that were found by the search operation.
        /// </summary>
        public Collection<Conversation> Conversations { get; internal set; }

        /// <summary>
        /// Gets a collection containing the HighlightTerms that were returned by the search operation.
        /// </summary>
        public Collection<HighlightTerm> HighlightTerms { get; internal set; }

        /// <summary>
        /// Gets the total count of conversations in view.
        /// </summary>
        public int? TotalCount { get; internal set; }

        /// <summary>
        /// Gets the indexed offset of the first conversation by the search operation.
        /// </summary>
        public int? IndexedOffset { get; internal set; }
    }
}
