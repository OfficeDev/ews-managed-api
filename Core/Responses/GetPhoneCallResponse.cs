// ---------------------------------------------------------------------------
// <copyright file="GetPhoneCallResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetPhoneCallResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a GetPhoneCall operation.
    /// </summary>
    internal sealed class GetPhoneCallResponse : ServiceResponse
    {
        private PhoneCall phoneCall;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetPhoneCallResponse"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetPhoneCallResponse(ExchangeService service)
            : base()
        {
            EwsUtilities.Assert(
                service != null,
                "GetPhoneCallResponse.ctor",
                "service is null");

            this.phoneCall = new PhoneCall(service);
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.PhoneCallInformation);
            this.phoneCall.LoadFromXml(reader, XmlNamespace.Messages, XmlElementNames.PhoneCallInformation);
            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.PhoneCallInformation);
        }

        /// <summary>
        /// Gets the phone call.
        /// </summary>
        internal PhoneCall PhoneCall
        {
            get
            {
                return this.phoneCall;
            }
        }     
    }
}