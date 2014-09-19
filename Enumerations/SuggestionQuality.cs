// ---------------------------------------------------------------------------
// <copyright file="SuggestionQuality.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SuggestionQuality enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the quality of an availability suggestion.
    /// </summary>
    public enum SuggestionQuality
    {
        /// <summary>
        /// The suggestion is excellent.
        /// </summary>
        Excellent,

        /// <summary>
        /// The suggestion is good.
        /// </summary>
        Good,

        /// <summary>
        /// The suggestion is fair.
        /// </summary>
        Fair,

        /// <summary>
        /// The suggestion is poor.
        /// </summary>
        Poor
    }
}
