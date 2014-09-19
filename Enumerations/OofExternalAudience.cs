// ---------------------------------------------------------------------------
// <copyright file="OofExternalAudience.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OofExternalAudience enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the external audience of an Out of Office notification.
    /// </summary>
    public enum OofExternalAudience
    {
        /// <summary>
        /// No external recipients should receive Out of Office notifications.
        /// </summary>
        None,

        /// <summary>
        /// Only recipients that are in the user's Contacts frolder should receive Out of Office notifications.
        /// </summary>
        Known,

        /// <summary>
        /// All recipients should receive Out of Office notifications.
        /// </summary>
        All
    }
}
