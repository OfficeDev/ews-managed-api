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
