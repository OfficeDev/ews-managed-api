// ---------------------------------------------------------------------------
// <copyright file="ICalendarActionProvider.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ICalendarActionProvider interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Interface defintion of a group of methods that are common to items that return CalendarActionResults
    /// </summary>
    internal interface ICalendarActionProvider
    {
        /// <summary>
        /// Implements the Accept method.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>A CalendarActionResults object containing the various items that were created or modified as a result of this operation.</returns>
        CalendarActionResults Accept(bool sendResponse);

        /// <summary>
        /// Implements the AcceptTentatively method.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>A CalendarActionResults object containing the various items that were created or modified as a result of this operation.</returns>
        CalendarActionResults AcceptTentatively(bool sendResponse);

        /// <summary>
        /// Implements the Decline method.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>A CalendarActionResults object containing the various items that were created or modified as a result of this operation.</returns>
        CalendarActionResults Decline(bool sendResponse);

        /// <summary>
        /// Implements the CreateAcceptMessage method.
        /// </summary>
        /// <param name="tentative">Indicates whether the new AcceptMeetingInvitationMessage should represent a Tentative accept response (as opposed to an Accept response).</param>
        /// <returns>A new AcceptMeetingInvitationMessage.</returns>
        AcceptMeetingInvitationMessage CreateAcceptMessage(bool tentative);

        /// <summary>
        /// Implements the DeclineMeetingInvitationMessage method.
        /// </summary>
        /// <returns>A new DeclineMeetingInvitationMessage.</returns>
        DeclineMeetingInvitationMessage CreateDeclineMessage();
    }
}
