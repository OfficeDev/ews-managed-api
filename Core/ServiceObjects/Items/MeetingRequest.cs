/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a meeting request that an attendee can accept or decline. Properties available on meeting requests are defined in the MeetingRequestSchema class.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.MeetingRequest)]
    public class MeetingRequest : MeetingMessage, ICalendarActionProvider
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingRequest"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal MeetingRequest(ItemAttachment parentAttachment)
            : base(parentAttachment)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingRequest"/> class.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal MeetingRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing meeting request and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting request.</param>
        /// <param name="id">The Id of the meeting request to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A MeetingRequest instance representing the meeting request corresponding to the specified Id.</returns>
        public static new MeetingRequest Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<MeetingRequest>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing meeting request and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting request.</param>
        /// <param name="id">The Id of the meeting request to bind to.</param>
        /// <returns>A MeetingRequest instance representing the meeting request corresponding to the specified Id.</returns>
        public static new MeetingRequest Bind(ExchangeService service, ItemId id)
        {
            return MeetingRequest.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return MeetingRequestSchema.Instance;
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Creates a local meeting acceptance message that can be customized and sent.
        /// </summary>
        /// <param name="tentative">Specifies whether the meeting will be tentatively accepted.</param>
        /// <returns>An AcceptMeetingInvitationMessage representing the meeting acceptance message. </returns>
        public AcceptMeetingInvitationMessage CreateAcceptMessage(bool tentative)
        {
            return new AcceptMeetingInvitationMessage(this, tentative);
        }

        /// <summary>
        /// Creates a local meeting declination message that can be customized and sent.
        /// </summary>
        /// <returns>A DeclineMeetingInvitation representing the meeting declination message. </returns>
        public DeclineMeetingInvitationMessage CreateDeclineMessage()
        {
            return new DeclineMeetingInvitationMessage(this);
        }

        /// <summary>
        /// Accepts the meeting. Calling this method results in a call to EWS. 
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public CalendarActionResults Accept(bool sendResponse)
        {
            return this.InternalAccept(false, sendResponse);
        }

        /// <summary>
        /// Tentatively accepts the meeting. Calling this method results in a call to EWS. 
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public CalendarActionResults AcceptTentatively(bool sendResponse)
        {
            return this.InternalAccept(true, sendResponse);
        }

        /// <summary>
        /// Accepts the meeting.
        /// </summary>
        /// <param name="tentative">True if tentative accept.</param>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        internal CalendarActionResults InternalAccept(bool tentative, bool sendResponse)
        {
            AcceptMeetingInvitationMessage accept = this.CreateAcceptMessage(tentative);

            if (sendResponse)
            {
                return accept.SendAndSaveCopy();
            }
            else
            {
                return accept.Save();
            }
        }

        /// <summary>
        /// Declines the meeting invitation. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="sendResponse">Indicates whether to send a response to the organizer.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public CalendarActionResults Decline(bool sendResponse)
        {
            DeclineMeetingInvitationMessage decline = this.CreateDeclineMessage();

            if (sendResponse)
            {
                return decline.SendAndSaveCopy();
            }
            else
            {
                return decline.Save();
            }
        }

        #region Properties

        /// <summary>
        /// Gets the type of this meeting request.
        /// </summary>
        public MeetingRequestType MeetingRequestType
        {
            get { return (MeetingRequestType)this.PropertyBag[MeetingRequestSchema.MeetingRequestType]; }
        }

        /// <summary>
        /// Gets the a value representing the intended free/busy status of the meeting.
        /// </summary>
        public LegacyFreeBusyStatus IntendedFreeBusyStatus
        {
            get { return (LegacyFreeBusyStatus)this.PropertyBag[MeetingRequestSchema.IntendedFreeBusyStatus]; }
        }

        /// <summary>
        /// Gets the change highlights of the meeting request.
        /// </summary>
        public ChangeHighlights ChangeHighlights
        {
            get { return (ChangeHighlights)this.PropertyBag[MeetingRequestSchema.ChangeHighlights]; }
        }

        /// <summary>
        /// Gets the Enhanced location object.
        /// </summary>
        public EnhancedLocation EnhancedLocation
        {
            get { return (EnhancedLocation)this.PropertyBag[MeetingRequestSchema.EnhancedLocation]; }
        }

        /// <summary>
        /// Gets the start time of the appointment.
        /// </summary>
        public DateTime Start
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.Start]; }
        }

        /// <summary>
        /// Gets the end time of the appointment.
        /// </summary>
        public DateTime End
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.End]; }
        }

        /// <summary>
        /// Gets the original start time of this appointment.
        /// </summary>
        public DateTime OriginalStart
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.OriginalStart]; }
        }

        /// <summary>
        /// Gets a value indicating whether this appointment is an all day event.
        /// </summary>
        public bool IsAllDayEvent
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsAllDayEvent]; }
        }

        /// <summary>
        /// Gets a value indicating the free/busy status of the owner of this appointment. 
        /// </summary>
        public LegacyFreeBusyStatus LegacyFreeBusyStatus
        {
            get { return (LegacyFreeBusyStatus)this.PropertyBag[AppointmentSchema.LegacyFreeBusyStatus]; }
        }

        /// <summary>
        /// Gets the location of this appointment.
        /// </summary>
        public string Location
        {
            get { return (string)this.PropertyBag[AppointmentSchema.Location]; }
        }

        /// <summary>
        /// Gets a text indicating when this appointment occurs. The text returned by When is localized using the Exchange Server culture or using the culture specified in the PreferredCulture property of the ExchangeService object this appointment is bound to.
        /// </summary>
        public string When
        {
            get { return (string)this.PropertyBag[AppointmentSchema.When]; }
        }

        /// <summary>
        /// Gets a value indicating whether the appointment is a meeting.
        /// </summary>
        public bool IsMeeting
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsMeeting]; }
        }

        /// <summary>
        ///  Gets a value indicating whether the appointment has been cancelled.
        /// </summary>
        public bool IsCancelled
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsCancelled]; }
        }

        /// <summary>
        ///  Gets a value indicating whether the appointment is recurring.
        /// </summary>
        public bool IsRecurring
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsRecurring]; }
        }

        /// <summary>
        ///  Gets a value indicating whether the meeting request has already been sent.
        /// </summary>
        public bool MeetingRequestWasSent
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.MeetingRequestWasSent]; }
        }

        /// <summary>
        /// Gets a value indicating the type of this appointment.
        /// </summary>
        public AppointmentType AppointmentType
        {
            get { return (AppointmentType)this.PropertyBag[AppointmentSchema.AppointmentType]; }
        }

        /// <summary>
        /// Gets a value indicating what was the last response of the user that loaded this meeting.
        /// </summary>
        public MeetingResponseType MyResponseType
        {
            get { return (MeetingResponseType)this.PropertyBag[AppointmentSchema.MyResponseType]; }
        }

        /// <summary>
        /// Gets the organizer of this meeting.
        /// </summary>
        public EmailAddress Organizer
        {
            get { return (EmailAddress)this.PropertyBag[AppointmentSchema.Organizer]; }
        }

        /// <summary>
        /// Gets a list of required attendees for this meeting.
        /// </summary>
        public AttendeeCollection RequiredAttendees
        {
            get { return (AttendeeCollection)this.PropertyBag[AppointmentSchema.RequiredAttendees]; }
        }

        /// <summary>
        /// Gets a list of optional attendeed for this meeting.
        /// </summary>
        public AttendeeCollection OptionalAttendees
        {
            get { return (AttendeeCollection)this.PropertyBag[AppointmentSchema.OptionalAttendees]; }
        }

        /// <summary>
        /// Gets a list of resources for this meeting.
        /// </summary>
        public AttendeeCollection Resources
        {
            get { return (AttendeeCollection)this.PropertyBag[AppointmentSchema.Resources]; }
        }

        /// <summary>
        /// Gets the number of calendar entries that conflict with this appointment in the authenticated user's calendar.
        /// </summary>
        public int ConflictingMeetingCount
        {
            get { return (int)this.PropertyBag[AppointmentSchema.ConflictingMeetingCount]; }
        }

        /// <summary>
        /// Gets the number of calendar entries that are adjacent to this appointment in the authenticated user's calendar.
        /// </summary>
        public int AdjacentMeetingCount
        {
            get { return (int)this.PropertyBag[AppointmentSchema.AdjacentMeetingCount]; }
        }

        /// <summary>
        /// Gets a list of meetings that conflict with this appointment in the authenticated user's calendar.
        /// </summary>
        public ItemCollection<Appointment> ConflictingMeetings
        {
            get { return (ItemCollection<Appointment>)this.PropertyBag[AppointmentSchema.ConflictingMeetings]; }
        }

        /// <summary>
        /// Gets a list of meetings that conflict with this appointment in the authenticated user's calendar.
        /// </summary>
        public ItemCollection<Appointment> AdjacentMeetings
        {
            get { return (ItemCollection<Appointment>)this.PropertyBag[AppointmentSchema.AdjacentMeetings]; }
        }

        /// <summary>
        /// Gets the duration of this appointment.
        /// </summary>
        public TimeSpan Duration
        {
            get { return (TimeSpan)this.PropertyBag[AppointmentSchema.Duration]; }
        }

        /// <summary>
        /// Gets the name of the time zone this appointment is defined in.
        /// </summary>
        public string TimeZone
        {
            get { return (string)this.PropertyBag[AppointmentSchema.TimeZone]; }
        }

        /// <summary>
        /// Gets the time when the attendee replied to the meeting request.
        /// </summary>
        public DateTime AppointmentReplyTime
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.AppointmentReplyTime]; }
        }

        /// <summary>
        /// Gets the sequence number of this appointment.
        /// </summary>
        public int AppointmentSequenceNumber
        {
            get { return (int)this.PropertyBag[AppointmentSchema.AppointmentSequenceNumber]; }
        }

        /// <summary>
        /// Gets the state of this appointment.
        /// </summary>
        public int AppointmentState
        {
            get { return (int)this.PropertyBag[AppointmentSchema.AppointmentState]; }
        }

        /// <summary>
        /// Gets the recurrence pattern for this meeting request.
        /// </summary>
        public Recurrence Recurrence
        {
            get { return (Recurrence)this.PropertyBag[AppointmentSchema.Recurrence]; }
        }

        /// <summary>
        /// Gets an OccurrenceInfo identifying the first occurrence of this meeting.
        /// </summary>
        public OccurrenceInfo FirstOccurrence
        {
            get { return (OccurrenceInfo)this.PropertyBag[AppointmentSchema.FirstOccurrence]; }
        }

        /// <summary>
        /// Gets an OccurrenceInfo identifying the last occurrence of this meeting.
        /// </summary>
        public OccurrenceInfo LastOccurrence
        {
            get { return (OccurrenceInfo)this.PropertyBag[AppointmentSchema.LastOccurrence]; }
        }

        /// <summary>
        /// Gets a list of modified occurrences for this meeting.
        /// </summary>
        public OccurrenceInfoCollection ModifiedOccurrences
        {
            get { return (OccurrenceInfoCollection)this.PropertyBag[AppointmentSchema.ModifiedOccurrences]; }
        }

        /// <summary>
        /// Gets a list of deleted occurrences for this meeting.
        /// </summary>
        public DeletedOccurrenceInfoCollection DeletedOccurrences
        {
            get { return (DeletedOccurrenceInfoCollection)this.PropertyBag[AppointmentSchema.DeletedOccurrences]; }
        }

        /// <summary>
        /// Gets time zone of the start property of this meeting request.
        /// </summary>
        public TimeZoneInfo StartTimeZone
        {
            get { return (TimeZoneInfo)this.PropertyBag[AppointmentSchema.StartTimeZone]; }
        }

        /// <summary>
        /// Gets time zone of the end property of this meeting request.
        /// </summary>
        public TimeZoneInfo EndTimeZone
        {
            get { return (TimeZoneInfo)this.PropertyBag[AppointmentSchema.EndTimeZone]; }
        }

        /// <summary>
        /// Gets the type of conferencing that will be used during the meeting.
        /// </summary>
        public int ConferenceType
        {
            get { return (int)this.PropertyBag[AppointmentSchema.ConferenceType]; }
        }

        /// <summary>
        /// Gets a value indicating whether new time proposals are allowed for attendees of this meeting.
        /// </summary>
        public bool AllowNewTimeProposal
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.AllowNewTimeProposal]; }
        }

        /// <summary>
        /// Gets a value indicating whether this is an online meeting.
        /// </summary>
        public bool IsOnlineMeeting
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsOnlineMeeting]; }
        }

        /// <summary>
        /// Gets the URL of the meeting workspace. A meeting workspace is a shared Web site for planning meetings and tracking results.
        /// </summary>
        public string MeetingWorkspaceUrl
        {
            get { return (string)this.PropertyBag[AppointmentSchema.MeetingWorkspaceUrl]; }
        }

        /// <summary>
        /// Gets the URL of the Microsoft NetShow online meeting.
        /// </summary>
        public string NetShowUrl
        {
            get { return (string)this.PropertyBag[AppointmentSchema.NetShowUrl]; }
        }

        #endregion
    }
}