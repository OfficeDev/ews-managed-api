// ---------------------------------------------------------------------------
// <copyright file="MeetingRequestType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingRequestType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the type of a meeting request.
    /// </summary>
    public enum MeetingRequestType
    {
        /// <summary>
        /// Undefined meeting request type.
        /// </summary>
        None,
        
        /// <summary>
        /// The meeting request is an update to the original meeting.
        /// </summary>
        FullUpdate,

        /// <summary>
        /// The meeting request is an information update.
        /// </summary>
        InformationalUpdate,

        /// <summary>
        /// The meeting request is for a new meeting.
        /// </summary>
        NewMeetingRequest,

        /// <summary>
        /// The meeting request is outdated.
        /// </summary>
        Outdated,

        /// <summary>
        /// The meeting update is a silent update to an existing meeting.
        /// </summary>
        SilentUpdate,

        /// <summary>
        /// The meeting update was forwarded to a delegate, and this copy is informational.
        /// </summary>
        PrincipalWantsCopy
    }
}
