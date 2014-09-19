// ---------------------------------------------------------------------------
// <copyright file="ContainmentMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContainmentMode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the containment mode for Contains search filters.
    /// </summary>
    public enum ContainmentMode
    {
        /// <summary>
        /// The comparison is between the full string and the constant. The property value and the supplied constant are precisely the same.
        /// </summary>
        FullString,

        /// <summary>
        /// The comparison is between the string prefix and the constant.
        /// </summary>
        Prefixed,

        /// <summary>
        /// The comparison is between a substring of the string and the constant.
        /// </summary>
        Substring,

        /// <summary>
        /// The comparison is between a prefix on individual words in the string and the constant.
        /// </summary>
        PrefixOnWords,

        /// <summary>
        /// The comparison is between an exact phrase in the string and the constant.
        /// </summary>
        ExactPhrase
    }
}
