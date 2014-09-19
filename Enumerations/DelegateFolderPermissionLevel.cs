// ---------------------------------------------------------------------------
// <copyright file="DelegateFolderPermissionLevel.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DelegateFolderPermissionLevel enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines a delegate user's permission level on a specific folder.
    /// </summary>
    public enum DelegateFolderPermissionLevel
    {
        /// <summary>
        /// The delegate has no permission.
        /// </summary>
        None,

        /// <summary>
        /// The delegate has Editor permissions.
        /// </summary>
        Editor,

        /// <summary>
        /// The delegate has Reviewer permissions.
        /// </summary>
        Reviewer,

        /// <summary>
        /// The delegate has Author permissions.
        /// </summary>
        Author,

        /// <summary>
        /// The delegate has custom permissions.
        /// </summary>
        Custom
    }
}
