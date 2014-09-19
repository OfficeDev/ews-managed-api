// ---------------------------------------------------------------------------
// <copyright file="FolderChange.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderChange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a change on a folder as returned by a synchronization operation.
    /// </summary>
    public sealed class FolderChange : Change
    {
        /// <summary>
        /// Initializes a new instance of FolderChange.
        /// </summary>
        internal FolderChange()
            : base()
        {
        }

        /// <summary>
        /// Creates a FolderId instance.
        /// </summary>
        /// <returns>A FolderId.</returns>
        internal override ServiceId CreateId()
        {
            return new FolderId();
        }

        /// <summary>
        /// Gets the folder the change applies to. Folder is null when ChangeType is equal to
        /// ChangeType.Delete. In that case, use the FolderId property to retrieve the Id of
        /// the folder that was deleted.
        /// </summary>
        public Folder Folder
        {
            get { return (Folder)this.ServiceObject; }
        }

        /// <summary>
        /// Gets the Id of the folder the change applies to.
        /// </summary>
        public FolderId FolderId
        {
            get { return (FolderId)this.Id; }
        }
    }
}
