// ---------------------------------------------------------------------------
// <copyright file="ChangeType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ChangeType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of change of a synchronization event.
    /// </summary>
    public enum ChangeType
    {
        /// <summary>
        /// An item or folder was created.
        /// </summary>
        Create,

        /// <summary>
        /// An item or folder was modified.
        /// </summary>
        Update,

        /// <summary>
        /// An item or folder was deleted.
        /// </summary>
        Delete,

        /// <summary>
        /// An item's IsRead flag was changed.
        /// </summary>
        ReadFlagChange,
    }
}
