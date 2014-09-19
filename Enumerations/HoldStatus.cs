// ---------------------------------------------------------------------------
// <copyright file="HoldStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the HoldStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the hold status.
    /// </summary>
    public enum HoldStatus
    {
        /// <summary>
        /// Not on hold
        /// </summary>
        NotOnHold,

        /// <summary>
        /// Placing/removing hold is in-progress
        /// </summary>
        Pending,

        /// <summary>
        /// On hold
        /// </summary>
        OnHold,

        /// <summary>
        /// Some mailboxes are on hold and some are not
        /// </summary>
        PartialHold,

        /// <summary>
        /// The hold operation failed
        /// </summary>
        Failed,
    }
}
