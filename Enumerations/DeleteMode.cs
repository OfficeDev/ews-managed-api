// ---------------------------------------------------------------------------
// <copyright file="DeleteMode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeleteMode enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents deletion modes.
    /// </summary>
    public enum DeleteMode
    {
        /// <summary>
        /// The item or folder will be permanently deleted.
        /// </summary>
        HardDelete,

        /// <summary>
        /// The item or folder will be moved to the dumpster. Items and folders in the dumpster can be recovered.
        /// </summary>
        SoftDelete,

        /// <summary>
        /// The item or folder will be moved to the mailbox' Deleted Items folder.
        /// </summary>
        MoveToDeletedItems
    }
}
