// ---------------------------------------------------------------------------
// <copyright file="ConnectionFailureCause.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConnectionFailureCause enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// The ConnectionFailureCause enumeration
    /// </summary>
    public enum ConnectionFailureCause
    {
        /// <summary>
        /// None
        /// </summary>
        None,

        /// <summary>
        /// UserBusy
        /// </summary>
        UserBusy,

        /// <summary>
        /// NoAnswer
        /// </summary>
        NoAnswer,

        /// <summary>
        /// Unavailable
        /// </summary>
        Unavailable,

        /// <summary>
        /// Other
        /// </summary>
        Other
    }
}