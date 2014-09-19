// ---------------------------------------------------------------------------
// <copyright file="DeclineMeetingInvitationMessage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeclineMeetingInvitationMessage class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a meeting declination message.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.DeclineItem, ReturnedByServer = false)]
    public sealed class DeclineMeetingInvitationMessage : CalendarResponseMessage<MeetingResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DeclineMeetingInvitationMessage"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        internal DeclineMeetingInvitationMessage(Item referenceItem)
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
    }
}
