// ---------------------------------------------------------------------------
// <copyright file="ResponseType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the types of response given to a meeting request.
    /// </summary>
    public enum MeetingResponseType
    {
        /// <summary>
        /// The response type is inknown.
        /// </summary>
        Unknown,

        /// <summary>
        /// There was no response. The authenticated is the organizer of the meeting.
        /// </summary>
        Organizer,

        /// <summary>
        /// The meeting was tentatively accepted.
        /// </summary>
        Tentative,

        /// <summary>
        /// The meeting was accepted.
        /// </summary>
        Accept,

        /// <summary>
        /// The meeting was declined.
        /// </summary>
        Decline,

        /// <summary>
        /// No response was received for the meeting.
        /// </summary>
        NoResponseReceived
    }
}
