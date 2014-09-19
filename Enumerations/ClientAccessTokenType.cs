// ---------------------------------------------------------------------------
// <copyright file="ClientAccessTokenType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of ClientAccessTokenType
    /// </summary>
    public enum ClientAccessTokenType
    {
        /// <summary>
        /// CallerIdentity
        /// </summary>
        CallerIdentity,

        /// <summary>
        /// ExtensionCallback.
        /// </summary>
        ExtensionCallback,

        /// <summary>
        /// ScopedToken.
        /// </summary>
        ScopedToken,
    }
}
