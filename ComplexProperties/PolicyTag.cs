// ---------------------------------------------------------------------------
// <copyright file="PolicyTag.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PolicyTag class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the policy tag of an item or folder.
    /// </summary>
    public sealed class PolicyTag : RetentionTagBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyTag"/> class.
        /// </summary>
        public PolicyTag()
            : base(XmlElementNames.PolicyTag)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyTag"/> class.
        /// </summary>
        /// <param name="isExplicit">Is explicit.</param>
        /// <param name="retentionId">Retention id.</param>
        public PolicyTag(bool isExplicit, Guid retentionId)
            : this()
        {
            this.IsExplicit = isExplicit;
            this.RetentionId = retentionId;
        }
    }
}
