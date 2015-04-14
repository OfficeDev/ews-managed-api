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
    using System.IO;
    using System.Text;

    using Microsoft.Exchange.WebServices.Data.Enumerations;

    /// <summary>
    /// Represents a GetPeopleInsights request.
    /// </summary>
    internal sealed class GetPeopleInsightsRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetPeopleInsightsRequest"/> class.
        /// </summary>
        /// <param name="service">The service</param>
        internal GetPeopleInsightsRequest(ExchangeService service)
            : base(service)
        {
            this.Emailaddresses = new List<string>();
        }

        /// <summary>
        /// Gets the collection of Emailaddress.
        /// </summary>
        internal List<string> Emailaddresses
        {
            get;
            set;
        }

        /// <summary>
        /// Validate the request
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            // TODO - Validate each emailaddress
            EwsUtilities.ValidateParamCollection(this.Emailaddresses, "EmailAddresses");
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetPeopleInsights;
        }

        /// <summary>
        /// Writes XML elements for GetPeopleInsights request
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.EmailAddresses);

            foreach (string emailAddress in this.Emailaddresses)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.String, emailAddress);
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetPeopleInsightsResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetPeopleInsightsResponse response = new GetPeopleInsightsResponse();
            response.LoadFromXml(reader, XmlElementNames.GetPeopleInsightsResponse);
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013_SP1;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetPeopleInsightsResponse Execute()
        {
            GetPeopleInsightsResponse serviceResponse = (GetPeopleInsightsResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}