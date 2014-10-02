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
