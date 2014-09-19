// ---------------------------------------------------------------------------
// <copyright file="HoldAction.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the HoldAction enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the hold action.
    /// </summary>
    public enum HoldAction
    {
        /// <summary>
        /// Create new hold
        /// </summary>
        Create,

        /// <summary>
        /// Update query associated with a hold
        /// </summary>
        Update,

        /// <summary>
        /// Release the hold
        /// </summary>
        Remove,
    }
}
