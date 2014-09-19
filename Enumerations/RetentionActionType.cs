// ---------------------------------------------------------------------------
// <copyright file="RetentionActionType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RetentionActionType enumeration.</summary>
//-----------------------------------------------------------------------

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the action of a retention policy tag.
    /// </summary>
    public enum RetentionActionType
    {
        /// <summary>
        /// Never tags (RetentionEnabled = false) do not have retention action in the FAI.
        /// </summary>
        None = 0,
        
        /// <summary>
        /// Expired items will be moved to the Deleted Items folder.
        /// </summary>
        MoveToDeletedItems = 1,

        /// <summary>
        /// Expired items will be moved to the organizational folder specified
        /// in the ExpirationDestination field.
        /// </summary>
        MoveToFolder = 2,

        /// <summary>
        /// Expired items will be soft deleted.
        /// </summary>
        DeleteAndAllowRecovery = 3,

        /// <summary>
        /// Expired items will be hard deleted.
        /// </summary>
        PermanentlyDelete = 4,

        /// <summary>
        /// Expired items will be tagged as expired.
        /// </summary>
        MarkAsPastRetentionLimit = 5,

        /// <summary>
        /// Expired items will be moved to the archive.
        /// </summary>
        MoveToArchive = 6,
    }
}
