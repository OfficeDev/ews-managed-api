// ---------------------------------------------------------------------------
// <copyright file="Sensitivity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Sensitivity enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the sensitivity of an item.
    /// </summary>
    public enum Sensitivity
    {
        /// <summary>
        /// The item has a normal sensitivity.
        /// </summary>
        Normal,

        /// <summary>
        /// The item is personal.
        /// </summary>
        Personal,

        /// <summary>
        /// The item is private.
        /// </summary>
        Private,

        /// <summary>
        /// The item is confidential.
        /// </summary>
        Confidential
    }
}
