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