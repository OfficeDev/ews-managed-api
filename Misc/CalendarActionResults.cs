#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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
