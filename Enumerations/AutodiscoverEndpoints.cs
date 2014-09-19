// ---------------------------------------------------------------------------
// <copyright file="AutodiscoverEndpoints.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AutodiscoverEndpoints enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the types of Autodiscover endpoints that are available.
    /// </summary>
    [Flags]
    internal enum AutodiscoverEndpoints
    {
        /// <summary>
        /// No endpoints available.
        /// </summary>
        None = 0,

        /// <summary>
        /// The "legacy" Autodiscover endpoint.
        /// </summary>
        Legacy = 1,

        /// <summary>
        /// The SOAP endpoint.
        /// </summary>
        Soap = 2,

        /// <summary>
        /// The WS-Security endpoint.
        /// </summary>
        WsSecurity = 4,

        /// <summary>
        /// The WS-Security/SymmetricKey endpoint.
        /// </summary>
        WSSecuritySymmetricKey = 8,

        /// <summary>
        /// The WS-Security/X509Cert endpoint.
        /// </summary>
        WSSecurityX509Cert = 16,

        /// <summary>
        /// The OAuth endpoint
        /// </summary>
        OAuth = 32,
    }
}
