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
    /// Represents a GetUserOofSettings request.
    /// </summary>
    internal sealed class GetUserOofSettingsRequest : SimpleServiceRequestBase
    {
        private string smtpAddress;

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetUserOofSettingsRequest;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            EwsUtilities.ValidateParam(this.SmtpAddress, "SmtpAddress");
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Mailbox);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Address, this.SmtpAddress);
            writer.WriteEndElement(); // Mailbox
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetUserOofSettingsResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetUserOofSettingsResponse serviceResponse = new GetUserOofSettingsResponse();

            serviceResponse.LoadFromXml(reader, XmlElementNames.ResponseMessage);

            if (serviceResponse.ErrorCode == ServiceError.NoError)
            {
                reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.OofSettings);

                serviceResponse.OofSettings = new OofSettings();
                serviceResponse.OofSettings.LoadFromXml(reader, reader.LocalName);

                serviceResponse.OofSettings.AllowExternalOof = reader.ReadElementValue<OofExternalAudience>(
                    XmlNamespace.Messages,
                    XmlElementNames.AllowExternalOof);
            }

            return serviceResponse;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserOofSettingsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetUserOofSettingsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetUserOofSettingsResponse Execute()
        {
            GetUserOofSettingsResponse serviceResponse = (GetUserOofSettingsResponse)this.InternalExecute();

            serviceResponse.ThrowIfNecessary();

            return serviceResponse;
        }

        /// <summary>
        /// Gets or sets the SMTP address.
        /// </summary>
        internal string SmtpAddress
        {
            get { return this.smtpAddress; }
            set { this.smtpAddress = value; }
        }
    }
}