// ---------------------------------------------------------------------------
// <copyright file="EffectiveRights.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EffectiveRights enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the effective user rights associated with an item or folder.
    /// </summary>
    [Flags]
    public enum EffectiveRights
    {
        /// <summary>
        /// The user has no acces right on the item or folder.
        /// </summary>
        None = 0,

        /// <summary>
        /// The user can create associated items (FAI)
        /// </summary>
        CreateAssociated = 1,

        /// <summary>
        /// The user can create items.
        /// </summary>
        CreateContents = 2,

        /// <summary>
        /// The user can create sub-folders.
        /// </summary>
        CreateHierarchy = 4,

        /// <summary>
        /// The user can delete items and/or folders.
        /// </summary>
        Delete = 8,

        /// <summary>
        /// The user can modify the properties of items and/or folders.
        /// </summary>
        Modify = 16,

        /// <summary>
        /// The user can read the contents of items.
        /// </summary>
        Read = 32,

        /// <summary>
        /// The user can view private items.
        /// </summary>
        ViewPrivateItems = 64
    }
}
