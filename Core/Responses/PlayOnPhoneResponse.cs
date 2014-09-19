// ---------------------------------------------------------------------------
// <copyright file="PlayOnPhoneResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PlayOnPhoneResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a PlayOnPhone operation
    /// </summary>
    internal sealed class PlayOnPhoneResponse : ServiceResponse
    {
        private PhoneCallId phoneCallId;

        /// <summary>
        /// Initializes a new instance of the <see cref="PlayOnPhoneResponse"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal PlayOnPhoneResponse(ExchangeService service)
            : base()
        {
            EwsUtilities.Assert(
                service != null,
                "PlayOnPhoneResponse.ctor",
                "service is null");

            this.phoneCallId = new PhoneCallId();
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.PhoneCallId);
            this.phoneCallId.LoadFromXml(reader, XmlNamespace.Messages, XmlElementNames.PhoneCallId);
            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.PhoneCallId);
        }

        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            this.phoneCallId.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.PhoneCallId), service);
        }

        /// <summary>
        /// Gets the Id of the phone call.
        /// </summary>
        internal PhoneCallId PhoneCallId
        {
            get
            {
                return this.phoneCallId;
            }
        }
    }
}
