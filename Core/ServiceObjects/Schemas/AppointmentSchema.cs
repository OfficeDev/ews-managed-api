// ---------------------------------------------------------------------------
// <copyright file="AppointmentSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AppointmentSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for appointment and meeting requests.
    /// </summary>
    [Schema]
    public class AppointmentSchema : ItemSchema
    {
        /// <summary>
        /// Field URIs for Appointment.
        /// </summary>
        private static class FieldUris
        {
            public const string Start = "calendar:Start";
            public const string End = "calendar:End";
            public const string OriginalStart = "calendar:OriginalStart";
            public const string IsAllDayEvent = "calendar:IsAllDayEvent";
            public const string LegacyFreeBusyStatus = "calendar:LegacyFreeBusyStatus";
            public const string Location = "calendar:Location";
            public const string When = "calendar:When";
            public const string IsMeeting = "calendar:IsMeeting";
            public const string IsCancelled = "calendar:IsCancelled";
            public const string IsRecurring = "calendar:IsRecurring";
            public const string MeetingRequestWasSent = "calendar:MeetingRequestWasSent";
            public const string IsResponseRequested = "calendar:IsResponseRequested";
            public const string CalendarItemType = "calendar:CalendarItemType";
            public const string MyResponseType = "calendar:MyResponseType";
            public const string Organizer = "calendar:Organizer";
            public const string RequiredAttendees = "calendar:RequiredAttendees";
            public const string OptionalAttendees = "calendar:OptionalAttendees";
            public const string Resources = "calendar:Resources";
            public const string ConflictingMeetingCount = "calendar:ConflictingMeetingCount";
            public const string AdjacentMeetingCount = "calendar:AdjacentMeetingCount";
            public const string ConflictingMeetings = "calendar:ConflictingMeetings";
            public const string AdjacentMeetings = "calendar:AdjacentMeetings";
            public const string Duration = "calendar:Duration";
            public const string TimeZone = "calendar:TimeZone";
            public const string AppointmentReplyTime = "calendar:AppointmentReplyTime";
            public const string AppointmentSequenceNumber = "calendar:AppointmentSequenceNumber";
            public const string AppointmentState = "calendar:AppointmentState";
            public const string Recurrence = "calendar:Recurrence";
            public const string FirstOccurrence = "calendar:FirstOccurrence";
            public const string LastOccurrence = "calendar:LastOccurrence";
            public const string ModifiedOccurrences = "calendar:ModifiedOccurrences";
            public const string DeletedOccurrences = "calendar:DeletedOccurrences";
            public const string MeetingTimeZone = "calendar:MeetingTimeZone";
            public const string StartTimeZone = "calendar:StartTimeZone";
            public const string EndTimeZone = "calendar:EndTimeZone";
            public const string ConferenceType = "calendar:ConferenceType";
            public const string AllowNewTimeProposal = "calendar:AllowNewTimeProposal";
            public const string IsOnlineMeeting = "calendar:IsOnlineMeeting";
            public const string MeetingWorkspaceUrl = "calendar:MeetingWorkspaceUrl";
            public const string NetShowUrl = "calendar:NetShowUrl";
            public const string Uid = "calendar:UID";
            public const string RecurrenceId = "calendar:RecurrenceId";
            public const string DateTimeStamp = "calendar:DateTimeStamp";
            public const string EnhancedLocation = "calendar:EnhancedLocation";
            public const string JoinOnlineMeetingUrl = "calendar:JoinOnlineMeetingUrl";
            public const string OnlineMeetingSettings = "calendar:OnlineMeetingSettings";
        }

        /// <summary>
        /// Defines the StartTimeZone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition StartTimeZone =
            new StartTimeZonePropertyDefinition(
                XmlElementNames.StartTimeZone,
                FieldUris.StartTimeZone,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the EndTimeZone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EndTimeZone =
            new TimeZonePropertyDefinition(
                XmlElementNames.EndTimeZone,
                FieldUris.EndTimeZone,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010);

        /// <summary>
        /// Defines the Start property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Start =
            new ScopedDateTimePropertyDefinition(
                XmlElementNames.Start,
                FieldUris.Start,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate(ExchangeVersion version)
                {
                    return AppointmentSchema.StartTimeZone;
                });

        /// <summary>
        /// Defines the End property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition End =
            new ScopedDateTimePropertyDefinition(
                XmlElementNames.End,
                FieldUris.End,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate(ExchangeVersion version)
                {
                    return version == ExchangeVersion.Exchange2007_SP1 ? AppointmentSchema.StartTimeZone : AppointmentSchema.EndTimeZone;
                });

        /// <summary>
        /// Defines the OriginalStart property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OriginalStart =
            new DateTimePropertyDefinition(
                XmlElementNames.OriginalStart,
                FieldUris.OriginalStart,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsAllDayEvent property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsAllDayEvent =
            new BoolPropertyDefinition(
                XmlElementNames.IsAllDayEvent,
                FieldUris.IsAllDayEvent,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the LegacyFreeBusyStatus property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition LegacyFreeBusyStatus =
            new GenericPropertyDefinition<LegacyFreeBusyStatus>(
                XmlElementNames.LegacyFreeBusyStatus,
                FieldUris.LegacyFreeBusyStatus,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Location =
            new StringPropertyDefinition(
                XmlElementNames.Location,
                FieldUris.Location,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the When property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition When =
            new StringPropertyDefinition(
                XmlElementNames.When,
                FieldUris.When,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsMeeting property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsMeeting =
            new BoolPropertyDefinition(
                XmlElementNames.IsMeeting,
                FieldUris.IsMeeting,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsCancelled property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsCancelled =
            new BoolPropertyDefinition(
                XmlElementNames.IsCancelled,
                FieldUris.IsCancelled,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsRecurring property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsRecurring =
            new BoolPropertyDefinition(
                XmlElementNames.IsRecurring,
                FieldUris.IsRecurring,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the MeetingRequestWasSent property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MeetingRequestWasSent =
            new BoolPropertyDefinition(
                XmlElementNames.MeetingRequestWasSent,
                FieldUris.MeetingRequestWasSent,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsResponseRequested property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsResponseRequested =
            new BoolPropertyDefinition(
                XmlElementNames.IsResponseRequested,
                FieldUris.IsResponseRequested,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the AppointmentType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentType =
            new GenericPropertyDefinition<AppointmentType>(
                XmlElementNames.CalendarItemType,
                FieldUris.CalendarItemType,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the MyResponseType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MyResponseType =
            new GenericPropertyDefinition<MeetingResponseType>(
                XmlElementNames.MyResponseType,
                FieldUris.MyResponseType,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Organizer property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Organizer =
            new ContainedPropertyDefinition<EmailAddress>(
                XmlElementNames.Organizer,
                FieldUris.Organizer,
                XmlElementNames.Mailbox,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddress(); });

        /// <summary>
        /// Defines the RequiredAttendees property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition RequiredAttendees =
            new ComplexPropertyDefinition<AttendeeCollection>(
                XmlElementNames.RequiredAttendees,
                FieldUris.RequiredAttendees,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new AttendeeCollection(); });

        /// <summary>
        /// Defines the OptionalAttendees property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OptionalAttendees =
            new ComplexPropertyDefinition<AttendeeCollection>(
                XmlElementNames.OptionalAttendees,
                FieldUris.OptionalAttendees,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new AttendeeCollection(); });

        /// <summary>
        /// Defines the Resources property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Resources =
            new ComplexPropertyDefinition<AttendeeCollection>(
                XmlElementNames.Resources,
                FieldUris.Resources,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new AttendeeCollection(); });

        /// <summary>
        /// Defines the ConflictingMeetingCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConflictingMeetingCount =
            new IntPropertyDefinition(
                XmlElementNames.ConflictingMeetingCount,
                FieldUris.ConflictingMeetingCount,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the AdjacentMeetingCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AdjacentMeetingCount =
            new IntPropertyDefinition(
                XmlElementNames.AdjacentMeetingCount,
                FieldUris.AdjacentMeetingCount,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ConflictingMeetings property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConflictingMeetings =
            new ComplexPropertyDefinition<ItemCollection<Appointment>>(
                XmlElementNames.ConflictingMeetings,
                FieldUris.ConflictingMeetings,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new ItemCollection<Appointment>(); });

        /// <summary>
        /// Defines the AdjacentMeetings property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AdjacentMeetings =
            new ComplexPropertyDefinition<ItemCollection<Appointment>>(
                XmlElementNames.AdjacentMeetings,
                FieldUris.AdjacentMeetings,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new ItemCollection<Appointment>(); });

        /// <summary>
        /// Defines the Duration property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Duration =
            new TimeSpanPropertyDefinition(
                XmlElementNames.Duration,
                FieldUris.Duration,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the TimeZone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition TimeZone =
            new StringPropertyDefinition(
                XmlElementNames.TimeZone,
                FieldUris.TimeZone,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the AppointmentReplyTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentReplyTime =
            new DateTimePropertyDefinition(
                XmlElementNames.AppointmentReplyTime,
                FieldUris.AppointmentReplyTime,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the AppointmentSequenceNumber property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentSequenceNumber =
            new IntPropertyDefinition(
                XmlElementNames.AppointmentSequenceNumber,
                FieldUris.AppointmentSequenceNumber,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the AppointmentState property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentState =
            new IntPropertyDefinition(
                XmlElementNames.AppointmentState,
                FieldUris.AppointmentState,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Recurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Recurrence =
            new RecurrencePropertyDefinition(
                XmlElementNames.Recurrence,
                FieldUris.Recurrence,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the FirstOccurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FirstOccurrence =
            new ComplexPropertyDefinition<OccurrenceInfo>(
                XmlElementNames.FirstOccurrence,
                FieldUris.FirstOccurrence,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new OccurrenceInfo(); });

        /// <summary>
        /// Defines the LastOccurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition LastOccurrence =
            new ComplexPropertyDefinition<OccurrenceInfo>(
                XmlElementNames.LastOccurrence,
                FieldUris.LastOccurrence,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new OccurrenceInfo(); });

        /// <summary>
        /// Defines the ModifiedOccurrences property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ModifiedOccurrences =
            new ComplexPropertyDefinition<OccurrenceInfoCollection>(
                XmlElementNames.ModifiedOccurrences,
                FieldUris.ModifiedOccurrences,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new OccurrenceInfoCollection(); });

        /// <summary>
        /// Defines the DeletedOccurrences property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DeletedOccurrences =
            new ComplexPropertyDefinition<DeletedOccurrenceInfoCollection>(
                XmlElementNames.DeletedOccurrences,
                FieldUris.DeletedOccurrences,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new DeletedOccurrenceInfoCollection(); });

        /// <summary>
        /// Defines the MeetingTimeZone property.
        /// </summary>
        internal static readonly PropertyDefinition MeetingTimeZone =
            new MeetingTimeZonePropertyDefinition(
                XmlElementNames.MeetingTimeZone,
                FieldUris.MeetingTimeZone,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ConferenceType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConferenceType =
            new IntPropertyDefinition(
                XmlElementNames.ConferenceType,
                FieldUris.ConferenceType,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the AllowNewTimeProposal property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AllowNewTimeProposal =
            new BoolPropertyDefinition(
                XmlElementNames.AllowNewTimeProposal,
                FieldUris.AllowNewTimeProposal,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsOnlineMeeting property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsOnlineMeeting =
            new BoolPropertyDefinition(
                XmlElementNames.IsOnlineMeeting,
                FieldUris.IsOnlineMeeting,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the MeetingWorkspaceUrl property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MeetingWorkspaceUrl =
            new StringPropertyDefinition(
                XmlElementNames.MeetingWorkspaceUrl,
                FieldUris.MeetingWorkspaceUrl,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the NetShowUrl property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition NetShowUrl =
            new StringPropertyDefinition(
                XmlElementNames.NetShowUrl,
                FieldUris.NetShowUrl,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the iCalendar Uid property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ICalUid =
            new StringPropertyDefinition(
                XmlElementNames.Uid,
                FieldUris.Uid,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the iCalendar RecurrenceId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ICalRecurrenceId =
            new DateTimePropertyDefinition(
                XmlElementNames.RecurrenceId,
                FieldUris.RecurrenceId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                true);  // isNullable

        /// <summary>
        /// Defines the iCalendar DateTimeStamp property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ICalDateTimeStamp =
            new DateTimePropertyDefinition(
                XmlElementNames.DateTimeStamp,
                FieldUris.DateTimeStamp,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                true);  // isNullable

        /// <summary>
        /// Enhanced Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EnhancedLocation = 
            new ComplexPropertyDefinition<EnhancedLocation>(
                XmlElementNames.EnhancedLocation,
                FieldUris.EnhancedLocation,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate() { return new EnhancedLocation(); });

        /// <summary>
        /// JoinOnlineMeetingUrl property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition JoinOnlineMeetingUrl =
            new StringPropertyDefinition(
                XmlElementNames.JoinOnlineMeetingUrl,
                FieldUris.JoinOnlineMeetingUrl,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// OnlineMeetingSettings property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OnlineMeetingSettings =
            new ComplexPropertyDefinition<OnlineMeetingSettings>(
                XmlElementNames.OnlineMeetingSettings,
                FieldUris.OnlineMeetingSettings,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate() { return new OnlineMeetingSettings(); });

        /// <summary>
        /// Instance of schema.
        /// </summary>
        /// <remarks>
        /// This must be after the declaration of property definitions.
        /// </remarks>
        internal static new readonly AppointmentSchema Instance = new AppointmentSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(ICalUid);
            this.RegisterProperty(ICalRecurrenceId);
            this.RegisterProperty(ICalDateTimeStamp);
            this.RegisterProperty(Start);
            this.RegisterProperty(End);
            this.RegisterProperty(OriginalStart);
            this.RegisterProperty(IsAllDayEvent);
            this.RegisterProperty(LegacyFreeBusyStatus);
            this.RegisterProperty(Location);
            this.RegisterProperty(When);
            this.RegisterProperty(IsMeeting);
            this.RegisterProperty(IsCancelled);
            this.RegisterProperty(IsRecurring);
            this.RegisterProperty(MeetingRequestWasSent);
            this.RegisterProperty(IsResponseRequested);
            this.RegisterProperty(AppointmentType);
            this.RegisterProperty(MyResponseType);
            this.RegisterProperty(Organizer);
            this.RegisterProperty(RequiredAttendees);
            this.RegisterProperty(OptionalAttendees);
            this.RegisterProperty(Resources);
            this.RegisterProperty(ConflictingMeetingCount);
            this.RegisterProperty(AdjacentMeetingCount);
            this.RegisterProperty(ConflictingMeetings);
            this.RegisterProperty(AdjacentMeetings);
            this.RegisterProperty(Duration);
            this.RegisterProperty(TimeZone);
            this.RegisterProperty(AppointmentReplyTime);
            this.RegisterProperty(AppointmentSequenceNumber);
            this.RegisterProperty(AppointmentState);
            this.RegisterProperty(Recurrence);
            this.RegisterProperty(FirstOccurrence);
            this.RegisterProperty(LastOccurrence);
            this.RegisterProperty(ModifiedOccurrences);
            this.RegisterProperty(DeletedOccurrences);
            this.RegisterInternalProperty(MeetingTimeZone);
            this.RegisterProperty(StartTimeZone);
            this.RegisterProperty(EndTimeZone);
            this.RegisterProperty(ConferenceType);
            this.RegisterProperty(AllowNewTimeProposal);
            this.RegisterProperty(IsOnlineMeeting);
            this.RegisterProperty(MeetingWorkspaceUrl);
            this.RegisterProperty(NetShowUrl);
            this.RegisterProperty(EnhancedLocation);
            this.RegisterProperty(JoinOnlineMeetingUrl);
            this.RegisterProperty(OnlineMeetingSettings);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AppointmentSchema"/> class.
        /// </summary>
        internal AppointmentSchema()
            : base()
        {
        }
    }
}