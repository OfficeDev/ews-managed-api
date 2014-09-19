// ---------------------------------------------------------------------------
// <copyright file="MemberStatus.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MemberStatus enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the status of group members.
    /// </summary>
    public enum MemberStatus
    {
        /// <summary>
        /// The member is unrecognized.
        /// </summary>
        Unrecognized,

        /// <summary>
        /// The member is normal.
        /// </summary>
        Normal,

        /// <summary>
        /// The member is demoted.
        /// </summary>
        Demoted
    }
}
