// ---------------------------------------------------------------------------
// <copyright file="FolderPermissionLevel.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderPermissionLevel enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    // TODO : Do we want to include more information about what those levels actually allow users to do?

    /// <summary>
    /// Defines permission levels for calendar folders.
    /// </summary>
    public enum FolderPermissionLevel
    {
        /// <summary>
        /// No permission is granted.
        /// </summary>
        None,

        /// <summary>
        /// The Owner level.
        /// </summary>
        Owner,

        /// <summary>
        /// The Publishing Editor level.
        /// </summary>
        PublishingEditor,

        /// <summary>
        /// The Editor level.
        /// </summary>
        Editor,

        /// <summary>
        /// The Pusnlishing Author level.
        /// </summary>
        PublishingAuthor,

        /// <summary>
        /// The Author level.
        /// </summary>
        Author,

        /// <summary>
        /// The Non-editing Author level.
        /// </summary>
        NoneditingAuthor,

        /// <summary>
        /// The Reviewer level.
        /// </summary>
        Reviewer,

        /// <summary>
        /// The Contributor level.
        /// </summary>
        Contributor,

        /// <summary>
        /// The Free/busy Time Only level. (Can only be applied to Calendar folders).
        /// </summary>
        FreeBusyTimeOnly,

        /// <summary>
        /// The Free/busy Time, Subject and Location level. (Can only be applied to Calendar folders).
        /// </summary>
        FreeBusyTimeAndSubjectAndLocation,

        /// <summary>
        /// The Custom level.
        /// </summary>
        Custom
    }
}
