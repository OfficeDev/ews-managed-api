// ---------------------------------------------------------------------------
// <copyright file="RetentionType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RetentionType enumeration.</summary>
//-----------------------------------------------------------------------

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the retention type.
    /// </summary>
    public enum RetentionType
    {
        /// <summary>
        /// Delete retention.
        /// </summary>
        Delete = 0,

        /// <summary>
        /// Archive retention.
        /// </summary>
        Archive = 1,
    }
}
