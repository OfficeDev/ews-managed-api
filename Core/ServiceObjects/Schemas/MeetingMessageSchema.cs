// ---------------------------------------------------------------------------
// <copyright file="MeetingMessageSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingMessageSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for meeting messages.
    /// </summary>
    [Schema]
    public class MeetingMessageSchema : EmailMessageSchema
    {
        /// <summary>
        /// Field URIs for MeetingMessage.
        /// </summary>
        private static class FieldUris
        {
            public const string AssociatedCalendarItemId = "meeting:AssociatedCalendarItemId";
            public const string IsDelegated = "meeting:IsDelegated";
            public const string IsOutOfDate = "meeting:IsOutOfDate";
            public const string HasBeenProcessed = "meeting:HasBeenProcessed";
            public const string ResponseType = "meeting:ResponseType";
            public const string IsOrganizer = "cal:IsOrganizer";
        }

        /// <summary>
        /// Defines the AssociatedAppointmentId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AssociatedAppointmentId =
            new ComplexPropertyDefinition<ItemId>(
                XmlElementNames.AssociatedCalendarItemId,
                FieldUris.AssociatedCalendarItemId,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new ItemId(); });

        /// <summary>
        /// Defines the IsDelegated property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsDelegated =
            new BoolPropertyDefinition(
                XmlElementNames.IsDelegated,
                FieldUris.IsDelegated,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsOutOfDate property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsOutOfDate =
            new BoolPropertyDefinition(
                XmlElementNames.IsOutOfDate,
                FieldUris.IsOutOfDate,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the HasBeenProcessed property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition HasBeenProcessed =
            new BoolPropertyDefinition(
                XmlElementNames.HasBeenProcessed,
                FieldUris.HasBeenProcessed,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ResponseType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ResponseType =
            new GenericPropertyDefinition<MeetingResponseType>(
                XmlElementNames.ResponseType,
                FieldUris.ResponseType,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the iCalendar Uid property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ICalUid =
            AppointmentSchema.ICalUid;

        /// <summary>
        /// Defines the iCalendar RecurrenceId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ICalRecurrenceId =
            AppointmentSchema.ICalRecurrenceId;

        /// <summary>
        /// Defines the iCalendar DateTimeStamp property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ICalDateTimeStamp =
            AppointmentSchema.ICalDateTimeStamp;

        /// <summary>
        /// Defines the IsOrganizer property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsOrganizer =
            new GenericPropertyDefinition<bool>(
                XmlElementNames.IsOrganizer,
                FieldUris.IsOrganizer,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        // This must be after the declaration of property definitions
        internal static new readonly MeetingMessageSchema Instance = new MeetingMessageSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(AssociatedAppointmentId);
            this.RegisterProperty(IsDelegated);
            this.RegisterProperty(IsOutOfDate);
            this.RegisterProperty(HasBeenProcessed);
            this.RegisterProperty(ResponseType);
            this.RegisterProperty(ICalUid);
            this.RegisterProperty(ICalRecurrenceId);
            this.RegisterProperty(ICalDateTimeStamp);
            this.RegisterProperty(IsOrganizer);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingMessageSchema"/> class.
        /// </summary>
        internal MeetingMessageSchema()
            : base()
        {
        }
    }
}