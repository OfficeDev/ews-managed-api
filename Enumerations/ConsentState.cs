// ---------------------------------------------------------------------------
// <copyright file="ConsentState.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data.Enumerations
{
    using System;

    /// <summary>
    /// The consent states enumeration
    /// </summary>
    [Serializable]
    public enum ConsentState
    {
        /// <summary>
        /// User has closed the consent page or has not responded yet.
        /// </summary>
        NotResponded = 0,

        /// <summary>
        /// User has requested to disable the extension.
        /// </summary>
        NotConsented = 1,

        /// <summary>
        /// User has requested to enable the extension.
        /// </summary>
        Consented = 2
    }
}
