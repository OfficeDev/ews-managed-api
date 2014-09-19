// ---------------------------------------------------------------------------
// <copyright file="ItemFlagStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemFlagStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the flag status of an Item.
    /// </summary>
    public enum ItemFlagStatus
    {
        /// <summary>
        /// Not Flagged.
        /// </summary>
        NotFlagged,

        /// <summary>
        /// Flagged.
        /// </summary>
        Flagged,

        /// <summary>
        /// Complete.
        /// </summary>
        Complete
    }
}