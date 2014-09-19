// ---------------------------------------------------------------------------
// <copyright file="PreviewItemBaseShape.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PreviewItemBaseShape enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Preview item base shape
    /// </summary>
    public enum PreviewItemBaseShape
    {
        /// <summary>
        /// Default (all properties required for showing preview by default)
        /// </summary>
        Default,

        /// <summary>
        /// Compact (only a set of core properties)
        /// </summary>
        Compact,
    }
}
