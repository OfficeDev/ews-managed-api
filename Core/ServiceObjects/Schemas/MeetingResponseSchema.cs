// ---------------------------------------------------------------------------
// <copyright file="MeetingResponseSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingResponseSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for meeting messages.
    /// </summary>
    [Schema]
    public class MeetingResponseSchema : MeetingMessageSchema
    {
        /// <summary>
        /// Field URIs for MeetingMessage.
        /// </summary>
        private static class FieldUris
        {
            public const string ProposedStart = "meeting:ProposedStart";
            public const string ProposedEnd = "meeting:ProposedEnd";
        }

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
        /// Defines the Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Location =
            AppointmentSchema.Location;

        /// <summary>
        /// Defines the AppointmentType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AppointmentType =
            AppointmentSchema.AppointmentType;

        /// <summary>
        /// Defines the Recurrence property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Recurrence =
            AppointmentSchema.Recurrence;

        /// <summary>
        /// Defines the Proposed Start property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ProposedStart =
            new ScopedDateTimePropertyDefinition(
                XmlElementNames.ProposedStart,
                FieldUris.ProposedStart,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate(ExchangeVersion version)
                {
                    return AppointmentSchema.StartTimeZone;
                });

        /// <summary>
        /// Defines the Proposed End property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ProposedEnd =
            new ScopedDateTimePropertyDefinition(
                XmlElementNames.ProposedEnd,
                FieldUris.ProposedEnd,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate(ExchangeVersion version)
                {
                    return AppointmentSchema.EndTimeZone;
                });

        /// <summary>
        /// Enhanced Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EnhancedLocation =
            AppointmentSchema.EnhancedLocation;

        // This must be after the declaration of property definitions
        internal static new readonly MeetingResponseSchema Instance = new MeetingResponseSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(Start);
            this.RegisterProperty(End);
            this.RegisterProperty(Location);
            this.RegisterProperty(Recurrence);
            this.RegisterProperty(AppointmentType);
            this.RegisterProperty(ProposedStart);
            this.RegisterProperty(ProposedEnd);
            this.RegisterProperty(EnhancedLocation);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingMessageSchema"/> class.
        /// </summary>
        internal MeetingResponseSchema()
            : base()
        {
        }
    }
}