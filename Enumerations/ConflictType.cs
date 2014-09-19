// ---------------------------------------------------------------------------
// <copyright file="ConflictType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConflictType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the conflict types that can be returned in meeting time suggestions.
    /// </summary>
    public enum ConflictType
    {
        /// <summary>
        /// There is a conflict with an indicidual attendee.
        /// </summary>
        IndividualAttendeeConflict,

        /// <summary>
        /// There is a conflict with at least one member of a group.
        /// </summary>
        GroupConflict,

        /// <summary>
        /// There is a conflict with at least one member of a group, but the group was too big for detailed information to be returned.
        /// </summary>
        GroupTooBigConflict,

        /// <summary>
        /// There is a conflict with an unresolvable attendee or an attendee that is not a user, group, or contact.
        /// </summary>
        UnknownAttendeeConflict
    }
}
