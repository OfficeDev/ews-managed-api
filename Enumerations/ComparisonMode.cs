// ---------------------------------------------------------------------------
// <copyright file="ComparisonMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ComparisonMode enumeration.</summary>
//-----------------------------------------------------------------------

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
