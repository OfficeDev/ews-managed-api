// ---------------------------------------------------------------------------
// <copyright file="SyncFolderHierarchyResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SyncFolderHierarchyResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a folder synchronization operation.
    /// </summary>
    public sealed class SyncFolderHierarchyResponse : SyncResponse<Folder, FolderChange>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncFolderHierarchyResponse"/> class.
        /// </summary>
        /// <param name="propertySet">Property set.</param>
        internal SyncFolderHierarchyResponse(PropertySet propertySet)
            : base(propertySet)
        {
        }

        /// <summary>
        /// Gets the name of the includes last in range XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetIncludesLastInRangeXmlElementName()
        {
            return XmlElementNames.IncludesLastFolderInRange;
        }

        /// <summary>
        /// Creates a folder change instance.
        /// </summary>
        /// <returns>FolderChange instance</returns>
        internal override FolderChange CreateChangeInstance()
        {
            return new FolderChange();
        }

        /// <summary>
        /// Gets the name of the change element.
        /// </summary>
        /// <returns>Change element name.</returns>
        internal override string GetChangeElementName()
        {
            return XmlElementNames.Folder;
        }

        /// <summary>
        /// Gets the name of the change id element.
        /// </summary>
        /// <returns>Change id element name.</returns>
        internal override string GetChangeIdElementName()
        {
            return XmlElementNames.FolderId;
        }

        /// <summary>
        /// Gets a value indicating whether this request returns full or summary properties.
        /// </summary>
        /// <value>
        /// <c>true</c> if summary properties only; otherwise, <c>false</c>.
        /// </value>
        internal override bool SummaryPropertiesOnly
        {
            get { return false; }
        }
    }
}
