// ---------------------------------------------------------------------------
// <copyright file="AcceptMeetingInvitationMessage.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AcceptMeetingInvitationMessage class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a meeting acceptance message.
    /// </summary>
    public sealed class AcceptMeetingInvitationMessage : CalendarResponseMessage<MeetingResponse>
    {
        private bool tentative;

        /// <summary>
        /// Initializes a new instance of the <see cref="AcceptMeetingInvitationMessage"/> class.
        /// </summary>
        /// <param name="referenceItem">The reference item.</param>
        /// <param name="tentative">if set to <c>true</c> accept invitation tentatively.</param>
        internal AcceptMeetingInvitationMessage(Item referenceItem, bool tentative)
            : base(referenceItem)
        {
            this.tentative = tentative;
        }

        /// <summary>
        /// This methods lets subclasses of ServiceObject override the default mechanism
        /// by which the XML element name associated with their type is retrieved.
        /// </summary>
        /// <returns>
        /// The XML element name associated with this type.
        /// If this method returns null or empty, the XML element name associated with this
        /// type is determined by the EwsObjectDefinition attribute that decorates the type,
        /// if present.
        /// </returns>
        /// <remarks>
        /// Item and folder classes that can be returned by EWS MUST rely on the EwsObjectDefinition
        /// attribute for XML element name determination.
        /// </remarks>
        internal override string GetXmlElementNameOverride()
        {
            if (this.tentative)
            {
                return XmlElementNames.TentativelyAcceptItem;
            }
            else
            {
                return XmlElementNames.AcceptItem;
            }
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
        /// Gets a value indicating whether the associated meeting is tentatively accepted.
        /// </summary>
        public bool Tentative
        {
            get { return this.tentative; }
        }
    }
}
