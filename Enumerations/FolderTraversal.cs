// ---------------------------------------------------------------------------
// <copyright file="FolderTraversal.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderTraversal enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the scope of FindFolders operations.
    /// </summary>
    public enum FolderTraversal
    {
        /// <summary>
        /// Only direct sub-folders are retrieved.
        /// </summary>
        Shallow,

        /// <summary>
        /// The entire hierarchy of sub-folders is retrieved.
        /// </summary>
        Deep,

        /// <summary>
        /// Only soft deleted folders are retrieved.
        /// </summary>
        SoftDeleted
    }
}
