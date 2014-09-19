// ---------------------------------------------------------------------------
// <copyright file="FolderPermissionReadAccess.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderPermissionReadAccess enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines a user's read access permission on items in a non-calendar folder.
    /// </summary>
    public enum FolderPermissionReadAccess
    {
        /// <summary>
        /// The user has no read access on the items in the folder.
        /// </summary>
        None,

        /// <summary>
        /// The user can read the start and end date and time of appointments. (Can only be applied to Calendar folders).
        /// </summary>
        TimeOnly,

        /// <summary>
        /// The user can read the start and end date and time, subject and location of appointments. (Can only be applied to Calendar folders).
        /// </summary>
        TimeAndSubjectAndLocation,

        /// <summary>
        /// The user has access to the full details of items.
        /// </summary>
        FullDetails
    }
}
