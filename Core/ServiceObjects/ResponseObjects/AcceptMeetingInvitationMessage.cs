#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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
