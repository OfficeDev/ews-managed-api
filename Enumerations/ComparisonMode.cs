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
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the way values are compared in search filters.
    /// </summary>
    public enum ComparisonMode
    {
        /// <summary>
        /// The comparison is exact.
        /// </summary>
        Exact,

        /// <summary>
        /// The comparison ignores casing.
        /// </summary>
        IgnoreCase,

        /// <summary>
        /// The comparison ignores spacing characters.
        /// </summary>
        IgnoreNonSpacingCharacters,

        /// <summary>
        /// The comparison ignores casing and spacing characters.
        /// </summary>
        IgnoreCaseAndNonSpacingCharacters

        // Although the following four values are defined in the EWS schema, they are useless
        // as they are all technically equivalent to Loose. We are not exposing those values
        // in this API. When we encounter one of these values on an existing search folder
        // restriction, we map it to IgnoreCaseAndNonSpacingCharacters.
        //
        // Loose,
        // LooseAndIgnoreCase,
        // LooseAndIgnoreNonSpace,
        // LooseAndIgnoreCaseAndIgnoreNonSpace
    }
}