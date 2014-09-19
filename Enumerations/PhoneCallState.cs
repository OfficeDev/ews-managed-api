// ---------------------------------------------------------------------------
// <copyright file="PhoneCallState.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhoneCallState enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// The PhoneCallState enumeration
    /// </summary>
    public enum PhoneCallState
    {
        /// <summary>
        /// Idle
        /// </summary>
        Idle,

        /// <summary>
        /// Connecting
        /// </summary>
        Connecting,

        /// <summary>
        /// Alerted
        /// </summary>
        Alerted,

        /// <summary>
        /// Connected
        /// </summary>
        Connected,

        /// <summary>
        /// Disconnected
        /// </summary>
        Disconnected,

        /// <summary>
        /// Incoming
        /// </summary>
        Incoming,

        /// <summary>
        /// Transferring
        /// </summary>
        Transferring,

        /// <summary>
        /// Forwarding
        /// </summary>
        Forwarding
    }
}
