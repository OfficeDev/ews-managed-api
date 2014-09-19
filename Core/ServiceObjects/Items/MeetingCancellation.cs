// ---------------------------------------------------------------------------
// <copyright file="MeetingCancellation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingCancellation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a meeting cancellation message. Properties available on meeting messages are defined in the MeetingMessageSchema class.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.MeetingCancellation)]
    public class MeetingCancellation : MeetingMessage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingCancellation"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal MeetingCancellation(ItemAttachment parentAttachment)
            : base(parentAttachment)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingCancellation"/> class.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal MeetingCancellation(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing meeting cancellation message and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting cancellation message.</param>
        /// <param name="id">The Id of the meeting cancellation message to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A MeetingCancellation instance representing the meeting cancellation message corresponding to the specified Id.</returns>
        public static new MeetingCancellation Bind(
            ExchangeService service,
            ItemId id, 
            PropertySet propertySet)
        {
            return service.BindToItem<MeetingCancellation>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing meeting cancellation message and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting cancellation message.</param>
        /// <param name="id">The Id of the meeting cancellation message to bind to.</param>
        /// <returns>A MeetingCancellation instance representing the meeting cancellation message corresponding to the specified Id.</returns>
        public static new MeetingCancellation Bind(ExchangeService service, ItemId id)
        {
            return MeetingCancellation.Bind(
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
            return MeetingCancellationSchema.Instance;
        }

        /// <summary>
        /// Removes the meeting associated with the cancellation message from the user's calendar.
        /// </summary>
        /// <returns>
        /// A CalendarActionResults object containing the various items that were created or modified as a
        /// results of this operation.
        /// </returns>
        public CalendarActionResults RemoveMeetingFromCalendar()
        {
            return new CalendarActionResults(new RemoveFromCalendar(this).InternalCreate(null, null));
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
        /// Gets the start time of the appointment.
        /// </summary>
        public DateTime Start
        {
            get { return (DateTime)this.PropertyBag[MeetingCancellationSchema.Start]; }
        }

        /// <summary>
        /// Gets the end time of the appointment.
        /// </summary>
        public DateTime End
        {
            get { return (DateTime)this.PropertyBag[MeetingCancellationSchema.End]; }
        }

        /// <summary>
        /// Gets the location of this appointment.
        /// </summary>
        public string Location
        {
            get { return (string)this.PropertyBag[MeetingCancellationSchema.Location]; }
        }

        /// <summary>
        /// Gets the recurrence pattern for this meeting request.
        /// </summary>
        public Recurrence Recurrence
        {
            get { return (Recurrence)this.PropertyBag[AppointmentSchema.Recurrence]; }
        }

        /// <summary>
        /// Gets the Enhanced location object.
        /// </summary>
        public EnhancedLocation EnhancedLocation
        {
            get { return (EnhancedLocation)this.PropertyBag[MeetingCancellationSchema.EnhancedLocation]; }
        }
    }
}
