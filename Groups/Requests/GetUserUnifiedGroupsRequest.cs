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

namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Exchange.WebServices.Data.Groups;

    /// <summary>
    /// Represents a request to a GetUserUnifiedGroupsRequest operation
    /// </summary>
    internal sealed class GetUserUnifiedGroupsRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserUnifiedGroupsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetUserUnifiedGroupsRequest(ExchangeService service) : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the RequestedUnifiedGroupsSet
        /// </summary>
        public IEnumerable<RequestedUnifiedGroupsSet> RequestedUnifiedGroupsSets { get; set; }

        /// <summary>
        /// Gets or sets the UserSmptAddress
        /// </summary>
        public string UserSmtpAddress { get; set; }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetUserUnifiedGroupsResponseMessage;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetUserUnifiedGroups;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetUserUnifiedGroupsResponse response = new GetUserUnifiedGroupsResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.RequestedGroupsSets);

            if (this.RequestedUnifiedGroupsSets != null)
            { 
                this.RequestedUnifiedGroupsSets.ForEach((unifiedGroupsSet) => unifiedGroupsSet.WriteToXml(writer, XmlElementNames.RequestedUnifiedGroupsSetItem));
            }

            writer.WriteEndElement();

            if (!string.IsNullOrEmpty(this.UserSmtpAddress))
            {
                writer.WriteElementValue(XmlNamespace.NotSpecified, XmlElementNames.UserSmtpAddress, this.UserSmtpAddress);
            }
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2015;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetUserUnifiedGroupsResponse Execute()
        {
            return (GetUserUnifiedGroupsResponse)this.InternalExecute();
        }
    }
}