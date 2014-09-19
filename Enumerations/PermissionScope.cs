// ---------------------------------------------------------------------------
// <copyright file="PermissionScope.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PermissionScope enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the scope of a user's permission on a folders.
    /// </summary>
    public enum PermissionScope
    {
        /// <summary>
        /// The user does not have the associated permission.
        /// </summary>
        None,

        /// <summary>
        /// The user has the associated permission on items that it owns.
        /// </summary>
        Owned,

        /// <summary>
        /// The user has the associated permission on all items.
        /// </summary>
        All
    }
}
