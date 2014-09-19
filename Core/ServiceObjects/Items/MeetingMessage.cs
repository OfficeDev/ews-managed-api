// ---------------------------------------------------------------------------
// <copyright file="MeetingMessage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingMessage class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.ComponentModel;

    /// <summary>
    /// Represents a meeting-related message. Properties available on meeting messages are defined in the MeetingMessageSchema class.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.MeetingMessage)]
    [EditorBrowsable(EditorBrowsableState.Never)]
    public class MeetingMessage : EmailMessage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingMessage"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal MeetingMessage(ItemAttachment parentAttachment)
            : base(parentAttachment)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingMessage"/> class.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal MeetingMessage(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing meeting message and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting message.</param>
        /// <param name="id">The Id of the meeting message to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A MeetingMessage instance representing the meeting message corresponding to the specified Id.</returns>
        public static new MeetingMessage Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<MeetingMessage>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing meeting message and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting message.</param>
        /// <param name="id">The Id of the meeting message to bind to.</param>
        /// <returns>A MeetingMessage instance representing the meeting message corresponding to the specified Id.</returns>
        public static new MeetingMessage Bind(ExchangeService service, ItemId id)
        {
            return MeetingMessage.Bind(
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
            return MeetingMessageSchema.Instance;
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        #region Properties

        /// <summary>
        /// Gets the Id of the appointment associated with the meeting message.
        /// </summary>
        public ItemId AssociatedAppointmentId
        {
            get { return (ItemId)this.PropertyBag[MeetingMessageSchema.AssociatedAppointmentId]; }
        }

        /// <summary>
        /// Gets a value indicating whether the meeting message is delegated.
        /// </summary>
        public bool IsDelegated
        {
            get { return (bool)this.PropertyBag[MeetingMessageSchema.IsDelegated]; }
        }

        /// <summary>
        /// Gets a value indicating whether the meeting message is out of date.
        /// </summary>
        public bool IsOutOfDate
        {
            get { return (bool)this.PropertyBag[MeetingMessageSchema.IsOutOfDate]; }
        }

        /// <summary>
        ///  Gets a value indicating whether the meeting message has been processed by Exchange (i.e. Exchange has noted
        ///  the arrival of a meeting request and has created the associated meeting item in the calendar).
        /// </summary>
        public bool HasBeenProcessed
        {
            get { return (bool)this.PropertyBag[MeetingMessageSchema.HasBeenProcessed]; }
        }

        /// <summary>
        /// Gets the isorganizer property for this meeting
        /// </summary>
        public bool? IsOrganizer
        {
            get { return (bool?)this.PropertyBag[MeetingMessageSchema.IsOrganizer]; }
        }

        /// <summary>
        /// Gets the type of response the meeting message represents.
        /// </summary>
        public MeetingResponseType ResponseType
        {
            get { return (MeetingResponseType)this.PropertyBag[MeetingMessageSchema.ResponseType]; }
        }

        /// <summary>
        /// Gets the ICalendar Uid.
        /// </summary>
        public string ICalUid
        {
            get { return (string)this.PropertyBag[MeetingMessageSchema.ICalUid]; }
        }

        /// <summary>
        /// Gets the ICalendar RecurrenceId.
        /// </summary>
        public DateTime? ICalRecurrenceId
        {
            get { return (DateTime?)this.PropertyBag[MeetingMessageSchema.ICalRecurrenceId]; }
        }

        /// <summary>
        /// Gets the ICalendar DateTimeStamp.
        /// </summary>
        public DateTime? ICalDateTimeStamp
        {
            get { return (DateTime?)this.PropertyBag[MeetingMessageSchema.ICalDateTimeStamp]; }
        }
        #endregion
    }
}
