// ---------------------------------------------------------------------------
// <copyright file="SearchFolderTraversal.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchFolderTraversal enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the scope of a search folder.
    /// </summary>
    public enum SearchFolderTraversal
    {
        /// <summary>
        /// Items belonging to the root folder are retrieved.
        /// </summary>
        Shallow,

        /// <summary>
        /// Items belonging to the root folder and its sub-folders are retrieved.
        /// </summary>
        Deep
    }
}
