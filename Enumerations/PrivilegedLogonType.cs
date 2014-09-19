// ---------------------------------------------------------------------------
// <copyright file="PrivilegedLogonType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PrivilegedLogonType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of PrivilegedLogonType.
    /// </summary>
    internal enum PrivilegedLogonType
    {
        /// <summary>
        /// Logon as Admin
        /// </summary>
        Admin,

        /// <summary>
        /// Logon as SystemService
        /// </summary>
        SystemService,
    }
}
