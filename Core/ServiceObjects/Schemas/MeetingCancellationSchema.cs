// ---------------------------------------------------------------------------
// <copyright file="MeetingCancellationSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingCancellationSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for meeting messages.
    /// </summary>
    [Schema]
    public class MeetingCancellationSchema : MeetingMessageSchema
    {
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
        /// Enhanced Location property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EnhancedLocation =
            AppointmentSchema.EnhancedLocation;

        // This must be after the declaration of property definitions
        internal static new readonly MeetingCancellationSchema Instance = new MeetingCancellationSchema();

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
            this.RegisterProperty(EnhancedLocation);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingMessageSchema"/> class.
        /// </summary>
        internal MeetingCancellationSchema()
            : base()
        {
        }
    }
}
