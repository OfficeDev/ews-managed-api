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
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for meeting requests.
    /// </summary>
    [Schema]
    public class MeetingRequestSchema : MeetingMessageSchema
    {
        /// <summary>
        /// Field URIs for MeetingRequest.
        /// </summary>
        private static class FieldUris
        {
            public const string MeetingRequestType = "meetingRequest:MeetingRequestType";
            public const string IntendedFreeBusyStatus = "meetingRequest:IntendedFreeBusyStatus";
            public const string ChangeHighlights = "meetingRequest:ChangeHighlights";
        }

        /// <summary>
        /// Defines the MeetingRequestType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MeetingRequestType =
            new GenericPropertyDefinition<MeetingRequestType>(
                XmlElementNames.MeetingRequestType,
                FieldUris.MeetingRequestType,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IntendedFreeBusyStatus property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IntendedFreeBusyStatus =
            new GenericPropertyDefinition<LegacyFreeBusyStatus>(
                XmlElementNames.IntendedFreeBusyStatus,
                FieldUris.IntendedFreeBusyStatus,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ChangeHighlights property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ChangeHighlights =
            new ComplexPropertyDefinition<ChangeHighlights>(
                XmlElementNames.ChangeHighlights,
                FieldUris.ChangeHighlights,
                ExchangeVersion.Exchange2013,
                delegate() { return new ChangeHighlights(); });

        /// <summary>
        /// Enhanced Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EnhancedLocation =
            AppointmentSchema.EnhancedLocation;

        /// <summary>
        /// Defines the Start property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Start =
            AppointmentSchema.Start;

        /// <summary>
        /// Defines the End property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition End =
            AppointmentSchema.End;

        /// <summary>
        /// Defines the OriginalStart property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OriginalStart =
            AppointmentSchema.OriginalStart;

        /// <summary>
        /// Defines the IsAllDayEvent property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsAllDayEvent =
            AppointmentSchema.IsAllDayEvent;

        /// <summary>
        /// Defines the LegacyFreeBusyStatus property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition LegacyFreeBusyStatus =
            AppointmentSchema.LegacyFreeBusyStatus;

        /// <summary>
        /// Defines the Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Location =
            AppointmentSchema.Location;

        /// <summary>
        /// Defines the When property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition When =
            AppointmentSchema.When;

        /// <summary>
        /// Defines the IsMeeting property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsMeeting =
            AppointmentSchema.IsMeeting;

        /// <summary>
        /// Defines the IsCancelled property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsCancelled =
            AppointmentSchema.IsCancelled;

        /// <summary>
        /// Defines the IsRecurring property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsRecurring =
            AppointmentSchema.IsRecurring;

        /// <summary>
        /// Defines the MeetingRequestWasSent property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MeetingRequestWasSent =
            AppointmentSchema.MeetingRequestWasSent;

        /// <summary>
        /// Defines the AppointmentType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentType =
            AppointmentSchema.AppointmentType;

        /// <summary>
        /// Defines the MyResponseType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MyResponseType =
            AppointmentSchema.MyResponseType;

        /// <summary>
        /// Defines the Organizer property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Organizer =
            AppointmentSchema.Organizer;

        /// <summary>
        /// Defines the RequiredAttendees property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition RequiredAttendees =
            AppointmentSchema.RequiredAttendees;

        /// <summary>
        /// Defines the OptionalAttendees property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OptionalAttendees =
            AppointmentSchema.OptionalAttendees;

        /// <summary>
        /// Defines the Resources property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Resources =
            AppointmentSchema.Resources;

        /// <summary>
        /// Defines the ConflictingMeetingCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConflictingMeetingCount =
            AppointmentSchema.ConflictingMeetingCount;

        /// <summary>
        /// Defines the AdjacentMeetingCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AdjacentMeetingCount =
            AppointmentSchema.AdjacentMeetingCount;

        /// <summary>
        /// Defines the ConflictingMeetings property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConflictingMeetings =
            AppointmentSchema.ConflictingMeetings;

        /// <summary>
        /// Defines the AdjacentMeetings property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AdjacentMeetings =
            AppointmentSchema.AdjacentMeetings;

        /// <summary>
        /// Defines the Duration property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Duration =
            AppointmentSchema.Duration;

        /// <summary>
        /// Defines the TimeZone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition TimeZone =
            AppointmentSchema.TimeZone;

        /// <summary>
        /// Defines the AppointmentReplyTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentReplyTime =
            AppointmentSchema.AppointmentReplyTime;

        /// <summary>
        /// Defines the AppointmentSequenceNumber property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentSequenceNumber =
            AppointmentSchema.AppointmentSequenceNumber;

        /// <summary>
        /// Defines the AppointmentState property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentState =
            AppointmentSchema.AppointmentState;

        /// <summary>
        /// Defines the Recurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Recurrence =
            AppointmentSchema.Recurrence;

        /// <summary>
        /// Defines the FirstOccurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FirstOccurrence =
            AppointmentSchema.FirstOccurrence;

        /// <summary>
        /// Defines the LastOccurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition LastOccurrence =
            AppointmentSchema.LastOccurrence;

        /// <summary>
        /// Defines the ModifiedOccurrences property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ModifiedOccurrences =
            AppointmentSchema.ModifiedOccurrences;

        /// <summary>
        /// Defines the DeletedOccurrences property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DeletedOccurrences =
            AppointmentSchema.DeletedOccurrences;

        /// <summary>
        /// Defines the MeetingTimeZone property.
        /// </summary>
        internal static readonly PropertyDefinition MeetingTimeZone =
            AppointmentSchema.MeetingTimeZone;

        /// <summary>
        /// Defines the StartTimeZone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition StartTimeZone =
            AppointmentSchema.StartTimeZone;

        /// <summary>
        /// Defines the EndTimeZone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EndTimeZone =
            AppointmentSchema.EndTimeZone;

        /// <summary>
        /// Defines the ConferenceType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConferenceType =
            AppointmentSchema.ConferenceType;

        /// <summary>
        /// Defines the AllowNewTimeProposal property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AllowNewTimeProposal =
            AppointmentSchema.AllowNewTimeProposal;

        /// <summary>
        /// Defines the IsOnlineMeeting property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsOnlineMeeting =
            AppointmentSchema.IsOnlineMeeting;

        /// <summary>
        /// Defines the MeetingWorkspaceUrl property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MeetingWorkspaceUrl =
            AppointmentSchema.MeetingWorkspaceUrl;

        /// <summary>
        /// Defines the NetShowUrl property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition NetShowUrl =
            AppointmentSchema.NetShowUrl;

        // This must be after the declaration of property definitions
        internal static new readonly MeetingRequestSchema Instance = new MeetingRequestSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(MeetingRequestType);
            this.RegisterProperty(IntendedFreeBusyStatus);
            this.RegisterProperty(ChangeHighlights);

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
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingRequestSchema"/> class.
        /// </summary>
        internal MeetingRequestSchema()
            : base()
        {
        }
    }
}