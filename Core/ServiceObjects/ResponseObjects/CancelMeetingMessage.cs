// ---------------------------------------------------------------------------
// <copyright file="CancelMeetingMessage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CancelMeetingMessage class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a meeting cancellation message.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.CancelCalendarItem, ReturnedByServer = false)]
    public sealed class CancelMeetingMessage : CalendarResponseMessageBase<MeetingCancellation>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CancelMeetingMessage"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal CancelMeetingMessage(Item referenceItem)
            : base(referenceItem)
        {
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
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return CancelMeetingMessageSchema.Instance;
        }

        #region Properties

        /// <summary>
        /// Gets or sets the body of the response.
        /// </summary>
        public MessageBody Body
        {
            get { return (MessageBody)this.PropertyBag[CancelMeetingMessageSchema.Body]; }
            set { this.PropertyBag[CancelMeetingMessageSchema.Body] = value; }
        }

        #endregion
    }
}
