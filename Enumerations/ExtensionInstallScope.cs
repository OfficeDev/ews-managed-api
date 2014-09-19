// ---------------------------------------------------------------------------
// <copyright file="ExtensionInstallScope.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ExtensionInstallScope enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of ExtensionInstallScope.
    /// </summary>
    public enum ExtensionInstallScope
    {
        /// <summary>
        /// Unassigned
        /// </summary>
        None = 0,

        /// <summary>
        /// User
        /// </summary>
        User = 1,

        /// <summary>
        /// Organization
        /// </summary>
        Organization = 2,

        /// <summary>
        /// Exchange Default
        /// </summary>
        Default = 3,
    }
}
