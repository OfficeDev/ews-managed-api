// ---------------------------------------------------------------------------
// <copyright file="OofState.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OofState enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines a user's Out of Office Assistant status.
    /// </summary>
    public enum OofState
    {
        /// <summary>
        /// The assistant is diabled.
        /// </summary>
        Disabled,

        /// <summary>
        /// The assistant is enabled.
        /// </summary>
        Enabled,

        /// <summary>
        /// The assistant is scheduled.
        /// </summary>
        Scheduled
    }
}
