/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents set hold on mailboxes parameters.
    /// </summary>
    public sealed class SetHoldOnMailboxesParameters
    {
        /// <summary>
        /// Action type
        /// </summary>
        public HoldAction ActionType { get; set; }

        /// <summary>
        /// Hold id
        /// </summary>
        public string HoldId { get; set; }

        /// <summary>
        /// Query
        /// </summary>
        public string Query { get; set; }

        /// <summary>
        /// Collection of mailboxes
        /// </summary>
        public string[] Mailboxes { get; set; }

        /// <summary>
        /// Query language
        /// </summary>
        public string Language { get; set; }

        /// <summary>
        /// In-place hold identity
        /// </summary>
        public string InPlaceHoldIdentity { get; set; }
    }
}