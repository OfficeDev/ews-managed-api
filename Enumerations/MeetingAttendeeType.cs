// ---------------------------------------------------------------------------
// <copyright file="MeetingAttendeeType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingAttendeeType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of a meeting attendee.
    /// </summary>
    public enum MeetingAttendeeType
    {
        /// <summary>
        /// The attendee is the organizer of the meeting.
        /// </summary>
        Organizer,

        /// <summary>
        /// The attendee is required.
        /// </summary>
        Required,

        /// <summary>
        /// The attendee is optional.
        /// </summary>
        Optional,

        /// <summary>
        /// The attendee is a room.
        /// </summary>
        Room,

        /// <summary>
        /// The attendee is a resource.
        /// </summary>
        Resource
    }
}
