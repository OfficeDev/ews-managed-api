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
    /// Represents a response to a meeting request. Properties available on meeting messages are defined in the MeetingMessageSchema class.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.MeetingResponse)]
    public class MeetingResponse : MeetingMessage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingResponse"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal MeetingResponse(ItemAttachment parentAttachment)
            : base(parentAttachment)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingResponse"/> class.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        internal MeetingResponse(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing meeting response and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting response.</param>
        /// <param name="id">The Id of the meeting response to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A MeetingResponse instance representing the meeting response corresponding to the specified Id.</returns>
        public static new MeetingResponse Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<MeetingResponse>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing meeting response and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the meeting response.</param>
        /// <param name="id">The Id of the meeting response to bind to.</param>
        /// <returns>A MeetingResponse instance representing the meeting response corresponding to the specified Id.</returns>
        public static new MeetingResponse Bind(ExchangeService service, ItemId id)
        {
            return MeetingResponse.Bind(
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
            return MeetingResponseSchema.Instance;
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
            get { return (DateTime)this.PropertyBag[MeetingResponseSchema.Start]; }
        }

        /// <summary>
        /// Gets the end time of the appointment.
        /// </summary>
        public DateTime End
        {
            get { return (DateTime)this.PropertyBag[MeetingResponseSchema.End]; }
        }

        /// <summary>
        /// Gets the location of this appointment.
        /// </summary>
        public string Location
        {
            get { return (string)this.PropertyBag[MeetingResponseSchema.Location]; }
        }
        
        /// <summary>
        /// Gets the recurrence pattern for this meeting request.
        /// </summary>
        public Recurrence Recurrence
        {
            get { return (Recurrence)this.PropertyBag[AppointmentSchema.Recurrence]; }
        }

        /// <summary>
        /// Gets the proposed start time of the appointment.
        /// </summary>
        public DateTime ProposedStart
        {
            get { return (DateTime)this.PropertyBag[MeetingResponseSchema.ProposedStart]; }
        }

        /// <summary>
        /// Gets the proposed end time of the appointment.
        /// </summary>
        public DateTime ProposedEnd
        {
            get { return (DateTime)this.PropertyBag[MeetingResponseSchema.ProposedEnd]; }
        }

        /// <summary>
        /// Gets the Enhanced location object.
        /// </summary>
        public EnhancedLocation EnhancedLocation
        {
            get { return (EnhancedLocation)this.PropertyBag[MeetingResponseSchema.EnhancedLocation]; }
        }
    }
}