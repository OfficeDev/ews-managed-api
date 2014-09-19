// ---------------------------------------------------------------------------
// <copyright file="ArchiveTag.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ArchiveTag class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the archive tag of an item or folder.
    /// </summary>
    public sealed class ArchiveTag : RetentionTagBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ArchiveTag"/> class.
        /// </summary>
        public ArchiveTag()
            : base(XmlElementNames.ArchiveTag)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ArchiveTag"/> class.
        /// </summary>
        /// <param name="isExplicit">Is explicit.</param>
        /// <param name="retentionId">Retention id.</param>
        public ArchiveTag(bool isExplicit, Guid retentionId)
            : this()
        {
            this.IsExplicit = isExplicit;
            this.RetentionId = retentionId;
        }
    }
}
