// ---------------------------------------------------------------------------
// <copyright file="Appointment.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Appointment class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents an appointment or a meeting. Properties available on appointments are defined in the AppointmentSchema class.
    /// </summary>
    [Attachable]
    [ServiceObjectDefinition(XmlElementNames.CalendarItem)]
    public class Appointment : Item, ICalendarActionProvider
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="Appointment"/>. To bind to an existing appointment, use Appointment.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService instance to which this appointmtnt is bound.</param>
        public Appointment(ExchangeService service)
            : base(service)
        {
            // If we're running against Exchange 2007, we need to explicitly preset
            // the StartTimeZone property since Exchange 2007 will otherwise scope
            // start and end to UTC.
            if (service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
            {
                this.StartTimeZone = service.TimeZone;
            }
        }

        /// <summary>
        /// Initializes a new instance of Appointment.
        /// </summary>
        /// <param name="parentAttachment">Parent attachment.</param>
        /// <param name="isNew">If true, attachment is new.</param>
        internal Appointment(ItemAttachment parentAttachment, bool isNew)
            : base(parentAttachment)
        {
            // If we're running against Exchange 2007, we need to explicitly preset
            // the StartTimeZone property since Exchange 2007 will otherwise scope
            // start and end to UTC.
            if (parentAttachment.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
            {
                if (isNew)
                {
                    this.StartTimeZone = parentAttachment.Service.TimeZone;
                }
            }
        }

        /// <summary>
        /// Binds to an existing appointment and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the appointment.</param>
        /// <param name="id">The Id of the appointment to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An Appointment instance representing the appointment corresponding to the specified Id.</returns>
        public static new Appointment Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<Appointment>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing appointment and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the appointment.</param>
        /// <param name="id">The Id of the appointment to bind to.</param>
        /// <returns>An Appointment instance representing the appointment corresponding to the specified Id.</returns>
        public static new Appointment Bind(ExchangeService service, ItemId id)
        {
            return Appointment.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an occurence of an existing appointment and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the appointment.</param>
        /// <param name="recurringMasterId">The Id of the recurring master that the index represents an occurrence of.</param>
        /// <param name="occurenceIndex">The index of the occurrence.</param>
        /// <returns>An Appointment instance representing the appointment occurence corresponding to the specified occurence index .</returns>
        public static Appointment BindToOccurrence(
            ExchangeService service,
            ItemId recurringMasterId,
            int occurenceIndex)
        {
            return Appointment.BindToOccurrence(
                service,
                recurringMasterId,
                occurenceIndex,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to an occurence of an existing appointment and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the appointment.</param>
        /// <param name="recurringMasterId">The Id of the recurring master that the index represents an occurrence of.</param>
        /// <param name="occurenceIndex">The index of the occurrence.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An Appointment instance representing the appointment occurence corresponding to the specified occurence index.</returns>
        public static Appointment BindToOccurrence(
            ExchangeService service,
            ItemId recurringMasterId,
            int occurenceIndex,
            PropertySet propertySet)
        {
            AppointmentOccurrenceId occurenceId = new AppointmentOccurrenceId(recurringMasterId.UniqueId, occurenceIndex);
            return Appointment.Bind(
                service,
                occurenceId,
                propertySet);
        }

        /// <summary>
        /// Binds to the master appointment of a recurring series and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the appointment.</param>
        /// <param name="occurrenceId">The Id of one of the occurrences in the series.</param>
        /// <returns>An Appointment instance representing the master appointment of the recurring series to which the specified occurrence belongs.</returns>
        public static Appointment BindToRecurringMaster(ExchangeService service, ItemId occurrenceId)
        {
            return Appointment.BindToRecurringMaster(
                service,
                occurrenceId,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Binds to the master appointment of a recurring series and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the appointment.</param>
        /// <param name="occurrenceId">The Id of one of the occurrences in the series.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An Appointment instance representing the master appointment of the recurring series to which the specified occurrence belongs.</returns>
        public static Appointment BindToRecurringMaster(
            ExchangeService service,
            ItemId occurrenceId,
            PropertySet propertySet)
            {
                RecurringAppointmentMasterId recurringMasterId = new RecurringAppointmentMasterId(occurrenceId.UniqueId);
                return Appointment.Bind(
                    service,
                    recurringMasterId,
                    propertySet);
            }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return AppointmentSchema.Instance;
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
        /// Gets a value indicating whether a time zone SOAP header should be emitted in a CreateItem
        /// or UpdateItem request so this item can be property saved or updated.
        /// </summary>
        /// <param name="isUpdateOperation">Indicates whether the operation being petrformed is an update operation.</param>
        /// <returns>
        ///     <c>true</c> if a time zone SOAP header should be emitted; otherwise, <c>false</c>.
        /// </returns>
        internal override bool GetIsTimeZoneHeaderRequired(bool isUpdateOperation)
        {
            if (isUpdateOperation)
            {
                return false;
            }
            else
            {
                bool isStartTimeZoneSetOrUpdated = this.PropertyBag.IsPropertyUpdated(AppointmentSchema.StartTimeZone);
                bool isEndTimeZoneSetOrUpdated = this.PropertyBag.IsPropertyUpdated(AppointmentSchema.EndTimeZone);

                if (isStartTimeZoneSetOrUpdated && isEndTimeZoneSetOrUpdated)
                {
                    // If both StartTimeZone and EndTimeZone have been set or updated and are the same as the service's
                    // time zone, we emit the time zone header and StartTimeZone and EndTimeZone are not emitted.
                    TimeZoneInfo startTimeZone;
                    TimeZoneInfo endTimeZone;

                    this.PropertyBag.TryGetProperty<TimeZoneInfo>(AppointmentSchema.StartTimeZone, out startTimeZone);
                    this.PropertyBag.TryGetProperty<TimeZoneInfo>(AppointmentSchema.EndTimeZone, out endTimeZone);

                    return startTimeZone == this.Service.TimeZone || endTimeZone == this.Service.TimeZone;
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// Determines whether properties defined with ScopedDateTimePropertyDefinition require custom time zone scoping.
        /// </summary>
        /// <returns>
        ///     <c>true</c> if this item type requires custom scoping for scoped date/time properties; otherwise, <c>false</c>.
        /// </returns>
        internal override bool GetIsCustomDateTimeScopingRequired()
        {
            return true;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            //  Make sure that if we're on the Exchange2007_SP1 schema version, if any of the following
            //  properties are set or updated:
            //      o   Start
            //      o   End
            //      o   IsAllDayEvent
            //      o   Recurrence
            //  ... then, we must send the MeetingTimeZone element (which is generated from StartTimeZone for
            //  Exchange2007_SP1 requests (see StartTimeZonePropertyDefinition.cs).  If the StartTimeZone isn't
            //  in the property bag, then throw, because clients must supply the proper time zone - either by
            //  loading it from a currently-existing appointment, or by setting it directly.  Otherwise, to dirty
            //  the StartTimeZone property, we just set it to its current value.
            if ((this.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1) &&
                !this.Service.Exchange2007CompatibilityMode)
            {
                if (this.PropertyBag.IsPropertyUpdated(AppointmentSchema.Start) ||
                    this.PropertyBag.IsPropertyUpdated(AppointmentSchema.End) ||
                    this.PropertyBag.IsPropertyUpdated(AppointmentSchema.IsAllDayEvent) ||
                    this.PropertyBag.IsPropertyUpdated(AppointmentSchema.Recurrence))
                {
                    //  If the property isn't in the property bag, throw....
                    if (!this.PropertyBag.Contains(AppointmentSchema.StartTimeZone))
                    {
                        throw new ServiceLocalException(Strings.StartTimeZoneRequired);
                    }

                    //  Otherwise, set the time zone to its current value to force it to be sent with the request.
                    this.StartTimeZone = this.StartTimeZone;
                }
            }
        }

        /// <summary>
        /// Creates a reply response to the organizer and/or attendees of the meeting.
        /// </summary>
        /// <param name="replyAll">Indicates whether the reply should go to the organizer only or to all the attendees.</param>
        /// <returns>A ResponseMessage representing the reply response that can subsequently be modified and sent.</returns>
        public ResponseMessage CreateReply(bool replyAll)
        {
            this.ThrowIfThisIsNew();

            return new ResponseMessage(
                this,
                replyAll ? ResponseMessageType.ReplyAll : ResponseMessageType.Reply);
        }

        /// <summary>
        /// Replies to the organizer and/or the attendees of the meeting. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the body of the meeting.</param>
        /// <param name="replyAll">Indicates whether the reply should go to the organizer only or to all the attendees.</param>
        public void Reply(MessageBody bodyPrefix, bool replyAll)
        {
            ResponseMessage responseMessage = this.CreateReply(replyAll);

            responseMessage.BodyPrefix = bodyPrefix;

            responseMessage.SendAndSaveCopy();
        }

        /// <summary>
        /// Creates a forward message from this appointment.
        /// </summary>
        /// <returns>A ResponseMessage representing the forward response that can subsequently be modified and sent.</returns>
        public ResponseMessage CreateForward()
        {
            this.ThrowIfThisIsNew();

            return new ResponseMessage(this, ResponseMessageType.Forward);
        }

        /// <summary>
        /// Forwards the appointment. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
        /// <param name="toRecipients">The recipients to forward the appointment to.</param>
        public void Forward(MessageBody bodyPrefix, params EmailAddress[] toRecipients)
        {
            this.Forward(bodyPrefix, (IEnumerable<EmailAddress>)toRecipients);
        }

        /// <summary>
        /// Forwards the appointment. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
        /// <param name="toRecipients">The recipients to forward the appointment to.</param>
        public void Forward(MessageBody bodyPrefix, IEnumerable<EmailAddress> toRecipients)
        {
            ResponseMessage responseMessage = this.CreateForward();

            responseMessage.BodyPrefix = bodyPrefix;
            responseMessage.ToRecipients.AddRange(toRecipients);

            responseMessage.SendAndSaveCopy();
        }

        /// <summary>
        /// Saves this appointment in the specified folder. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save this appointment.</param>
        /// <param name="sendInvitationsMode">Specifies if and how invitations should be sent if this appointment is a meeting.</param>
        public void Save(WellKnownFolderName destinationFolderName, SendInvitationsMode sendInvitationsMode)
        {
            this.InternalCreate(
                new FolderId(destinationFolderName),
                null,
                sendInvitationsMode);
        }

        /// <summary>
        /// Saves this appointment in the specified folder. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save this appointment.</param>
        /// <param name="sendInvitationsMode">Specifies if and how invitations should be sent if this appointment is a meeting.</param>
        public void Save(FolderId destinationFolderId, SendInvitationsMode sendInvitationsMode)
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            this.InternalCreate(
                destinationFolderId,
                null,
                sendInvitationsMode);
        }

        /// <summary>
        /// Saves this appointment in the Calendar folder. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added.
        /// </summary>
        /// <param name="sendInvitationsMode">Specifies if and how invitations should be sent if this appointment is a meeting.</param>
        public void Save(SendInvitationsMode sendInvitationsMode)
        {
            this.InternalCreate(
                null,
                null,
                sendInvitationsMode);
        }

        /// <summary>
        /// Applies the local changes that have been made to this appointment. Calling this method results in at least one call to EWS.
        /// Mutliple calls to EWS might be made if attachments have been added or removed.
        /// </summary>
        /// <param name="conflictResolutionMode">Specifies how conflicts should be resolved.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Specifies if and how invitations or cancellations should be sent if this appointment is a meeting.</param>
        public void Update(
            ConflictResolutionMode conflictResolutionMode,
            SendInvitationsOrCancellationsMode sendInvitationsOrCancellationsMode)
        {
            this.InternalUpdate(
                null,
                conflictResolutionMode,
                null,
                sendInvitationsOrCancellationsMode);
        }

        /// <summary>
        /// Deletes this appointment. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Specifies if and how cancellations should be sent if this appointment is a meeting.</param>
        public void Delete(DeleteMode deleteMode, SendCancellationsMode sendCancellationsMode)
        {
            this.InternalDelete(
                deleteMode,
                sendCancellationsMode,
                null);
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
        /// Creates a local meeting cancellation message that can be customized and sent.
        /// </summary>
        /// <returns>A CancelMeetingMessage representing the meeting cancellation message. </returns>
        public CancelMeetingMessage CreateCancelMeetingMessage()
        {
            return new CancelMeetingMessage(this);
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
        /// Cancels the meeting and sends cancellation messages to all attendees. Calling this method results in a call to EWS.
        /// </summary>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public CalendarActionResults CancelMeeting()
        {
            return this.CreateCancelMeetingMessage().SendAndSaveCopy();
        }

        /// <summary>
        /// Cancels the meeting and sends cancellation messages to all attendees. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="cancellationMessageText">Cancellation message text sent to all attendees.</param>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public CalendarActionResults CancelMeeting(string cancellationMessageText)
        {
            CancelMeetingMessage cancelMsg = this.CreateCancelMeetingMessage();
            cancelMsg.Body = cancellationMessageText;
            return cancelMsg.SendAndSaveCopy();
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

        /// <summary>
        /// Gets the default setting for sending cancellations on Delete.
        /// </summary>
        /// <returns>If Delete() is called on Appointment, we want to send cancellations and save a copy.</returns>
        internal override SendCancellationsMode? DefaultSendCancellationsMode
        {
            get { return SendCancellationsMode.SendToAllAndSaveCopy; }
        }

        /// <summary>
        /// Gets the default settings for sending invitations on Save.
        /// </summary>
        internal override SendInvitationsMode? DefaultSendInvitationsMode
        {
            get { return SendInvitationsMode.SendToAllAndSaveCopy; }
        }

        /// <summary>
        /// Gets the default settings for sending invitations or cancellations on Update.
        /// </summary>
        internal override SendInvitationsOrCancellationsMode? DefaultSendInvitationsOrCancellationsMode
        {
            get { return SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy; }
        }

        #region Properties
        /// <summary>
        /// Gets or sets the start time of the appointment.
        /// </summary>
        public DateTime Start
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.Start]; }
            set { this.PropertyBag[AppointmentSchema.Start] = value; }
        }

        /// <summary>
        /// Gets or sets the end time of the appointment.
        /// </summary>
        public DateTime End
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.End]; }
            set { this.PropertyBag[AppointmentSchema.End] = value; }
        }

        /// <summary>
        /// Gets the original start time of this appointment.
        /// </summary>
        public DateTime OriginalStart
        {
            get { return (DateTime)this.PropertyBag[AppointmentSchema.OriginalStart]; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this appointment is an all day event.
        /// </summary>
        public bool IsAllDayEvent
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsAllDayEvent]; }
            set { this.PropertyBag[AppointmentSchema.IsAllDayEvent] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating the free/busy status of the owner of this appointment. 
        /// </summary>
        public LegacyFreeBusyStatus LegacyFreeBusyStatus
        {
            get { return (LegacyFreeBusyStatus)this.PropertyBag[AppointmentSchema.LegacyFreeBusyStatus]; }
            set { this.PropertyBag[AppointmentSchema.LegacyFreeBusyStatus] = value; }
        }

        /// <summary>
        /// Gets or sets the location of this appointment.
        /// </summary>
        public string Location
        {
            get { return (string)this.PropertyBag[AppointmentSchema.Location]; }
            set { this.PropertyBag[AppointmentSchema.Location] = value; }
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
        /// Gets a value indicating whether the appointment has been cancelled.
        /// </summary>
        public bool IsCancelled
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsCancelled]; }
        }

        /// <summary>
        /// Gets a value indicating whether the appointment is recurring.
        /// </summary>
        public bool IsRecurring
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsRecurring]; }
        }

        /// <summary>
        /// Gets a value indicating whether the meeting request has already been sent.
        /// </summary>
        public bool MeetingRequestWasSent
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.MeetingRequestWasSent]; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether responses are requested when invitations are sent for this meeting.
        /// </summary>
        public bool IsResponseRequested
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsResponseRequested]; }
            set { this.PropertyBag[AppointmentSchema.IsResponseRequested] = value; }
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
        /// Gets the organizer of this meeting. The Organizer property is read-only and is only relevant for attendees.
        /// The organizer of a meeting is automatically set to the user that created the meeting.
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
        /// Gets or sets the recurrence pattern for this appointment. Available recurrence pattern classes include
        /// Recurrence.DailyPattern, Recurrence.MonthlyPattern and Recurrence.YearlyPattern.
        /// </summary>
        public Recurrence Recurrence
        {
            get
            {
                return (Recurrence)this.PropertyBag[AppointmentSchema.Recurrence];
            }

            set
            {
                if (value != null)
                {
                    if (value.IsRegenerationPattern)
                    {
                        throw new ServiceLocalException(Strings.RegenerationPatternsOnlyValidForTasks);
                    }
                }

                this.PropertyBag[AppointmentSchema.Recurrence] = value;
            }
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
        /// Gets or sets time zone of the start property of this appointment.
        /// </summary>
        public TimeZoneInfo StartTimeZone
        {
            get { return (TimeZoneInfo)this.PropertyBag[AppointmentSchema.StartTimeZone]; }
            set { this.PropertyBag[AppointmentSchema.StartTimeZone] = value; }
        }

        /// <summary>
        /// Gets or sets time zone of the end property of this appointment.
        /// </summary>
        public TimeZoneInfo EndTimeZone
        {
            get { return (TimeZoneInfo)this.PropertyBag[AppointmentSchema.EndTimeZone]; }
            set { this.PropertyBag[AppointmentSchema.EndTimeZone] = value; }
        }

        /// <summary>
        /// Gets or sets the type of conferencing that will be used during the meeting.
        /// </summary>
        public int ConferenceType
        {
            get { return (int)this.PropertyBag[AppointmentSchema.ConferenceType]; }
            set { this.PropertyBag[AppointmentSchema.ConferenceType] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether new time proposals are allowed for attendees of this meeting.
        /// </summary>
        public bool AllowNewTimeProposal
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.AllowNewTimeProposal]; }
            set { this.PropertyBag[AppointmentSchema.AllowNewTimeProposal] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this is an online meeting.
        /// </summary>
        public bool IsOnlineMeeting
        {
            get { return (bool)this.PropertyBag[AppointmentSchema.IsOnlineMeeting]; }
            set { this.PropertyBag[AppointmentSchema.IsOnlineMeeting] = value; }
        }

        /// <summary>
        /// Gets or sets the URL of the meeting workspace. A meeting workspace is a shared Web site for planning meetings and tracking results.
        /// </summary>
        public string MeetingWorkspaceUrl
        {
            get { return (string)this.PropertyBag[AppointmentSchema.MeetingWorkspaceUrl]; }
            set { this.PropertyBag[AppointmentSchema.MeetingWorkspaceUrl] = value; }
        }

        /// <summary>
        /// Gets or sets the URL of the Microsoft NetShow online meeting.
        /// </summary>
        public string NetShowUrl
        {
            get { return (string)this.PropertyBag[AppointmentSchema.NetShowUrl]; }
            set { this.PropertyBag[AppointmentSchema.NetShowUrl] = value; }
        }
        
        /// <summary>
        /// Gets or sets the ICalendar Uid.
        /// </summary>
        public string ICalUid
        {
            get { return (string)this.PropertyBag[AppointmentSchema.ICalUid]; }
            set { this.PropertyBag[AppointmentSchema.ICalUid] = value; }
        }

        /// <summary>
        /// Gets the ICalendar RecurrenceId.
        /// </summary>
        public DateTime? ICalRecurrenceId
        {
            get { return (DateTime?)this.PropertyBag[AppointmentSchema.ICalRecurrenceId]; }
        }

        /// <summary>
        /// Gets the ICalendar DateTimeStamp.
        /// </summary>
        public DateTime? ICalDateTimeStamp
        {
            get { return (DateTime?)this.PropertyBag[AppointmentSchema.ICalDateTimeStamp]; }
        }

        /// <summary>
        /// Gets or sets the Enhanced location object.
        /// </summary>
        public EnhancedLocation EnhancedLocation
        {
            get { return (EnhancedLocation)this.PropertyBag[AppointmentSchema.EnhancedLocation]; }
            set { this.PropertyBag[AppointmentSchema.EnhancedLocation] = value; }
        }

        /// <summary>
        /// Gets the Url for joining an online meeting
        /// </summary>
        public string JoinOnlineMeetingUrl
        {
            get { return (string)this.PropertyBag[AppointmentSchema.JoinOnlineMeetingUrl]; }
        }

        /// <summary>
        /// Gets the Online Meeting Settings
        /// </summary>
        public OnlineMeetingSettings OnlineMeetingSettings
        {
            get { return (OnlineMeetingSettings)this.PropertyBag[AppointmentSchema.OnlineMeetingSettings]; }
        }
        #endregion
    }
}