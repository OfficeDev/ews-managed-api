// ---------------------------------------------------------------------------
// <copyright file="CalendarActionResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CalendarActionResults class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the results of an action performed on a calendar item or meeting message,
    /// such as accepting, tentatively accepting or declining a meeting request.
    /// </summary>
    public sealed class CalendarActionResults
    {
        private Appointment appointment;
        private MeetingRequest meetingRequest;
        private MeetingResponse meetingResponse;
        private MeetingCancellation meetingCancellation;

        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarActionResults"/> class.
        /// </summary>
        /// <param name="items">Collection of items that were created or modified as a result of a calendar action.</param>
        internal CalendarActionResults(IEnumerable<Item> items)
        {
            this.appointment = EwsUtilities.FindFirstItemOfType<Appointment>(items);
            this.meetingRequest = EwsUtilities.FindFirstItemOfType<MeetingRequest>(items);
            this.meetingResponse = EwsUtilities.FindFirstItemOfType<MeetingResponse>(items);
            this.meetingCancellation = EwsUtilities.FindFirstItemOfType<MeetingCancellation>(items);
        }

        /// <summary>
        /// Gets the meeting that was accepted, tentatively accepted or declined.
        /// </summary>
        /// <remarks>
        /// When a meeting is accepted or tentatively accepted via an Appointment object,
        /// EWS recreates the meeting, and Appointment represents that new version.
        /// When a meeting is accepted or tentatively accepted via a MeetingRequest object,
        /// EWS creates an associated meeting in the attendee's calendar and Appointment
        /// represents that meeting.
        /// When declining a meeting via an Appointment object, EWS moves the appointment to
        /// the attendee's Deleted Items folder and Appointment represents that moved copy.
        /// When declining a meeting via a MeetingRequest object, EWS creates an associated
        /// meeting in the attendee's Deleted Items folder, and Appointment represents that
        /// meeting.
        /// When a meeting is declined via either an Appointment or a MeetingRequest object
        /// from the Deleted Items folder, Appointment is null.
        /// </remarks>
        public Appointment Appointment
        {
            get { return this.appointment; }
        }

        /// <summary>
        /// Gets the meeting request that was moved to the Deleted Items folder as a result
        /// of an attendee accepting, tentatively accepting or declining a meeting request.
        /// If the meeting request is accepted, tentatively accepted or declined from the
        /// Deleted Items folder, it is permanently deleted and MeetingRequest is null.
        /// </summary>
        public MeetingRequest MeetingRequest
        {
            get { return this.meetingRequest; }
        }

        /// <summary>
        /// Gets the copy of the response that is sent to the organizer of a meeting when
        /// the meeting is accepted, tentatively accepted or declined by an attendee.
        /// MeetingResponse is null if the attendee chose not to send a response.
        /// </summary>
        public MeetingResponse MeetingResponse
        {
            get { return this.meetingResponse; }
        }

        /// <summary>
        /// Gets the copy of the meeting cancellation message sent by the organizer to the
        /// attendees of a meeting when the meeting is cancelled.
        /// </summary>
        public MeetingCancellation MeetingCancellation
        {
            get { return this.meetingCancellation; }
        }
    }
}
