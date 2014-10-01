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
// <summary>Defines the GetServerTimeZonesRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetServerTimeZones request.
    /// </summary>
    internal class GetServerTimeZonesRequest : MultiResponseServiceRequest<GetServerTimeZonesResponse>
    {
        private IEnumerable<string> ids;

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.ids != null)
            {
                EwsUtilities.ValidateParamCollection(this.ids, "Ids");
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GetServerTimeZonesRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetServerTimeZonesRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override GetServerTimeZonesResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new GetServerTimeZonesResponse();
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.GetServerTimeZonesResponseMessage;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return 1;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetServerTimeZones;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetServerTimeZonesResponse;
        }

        /// <summary>
        /// Gets the minimum server version required to process this request.
        /// </summary>
        /// <returns>Exchange server version.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.Ids != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Ids);

                foreach (string id in this.ids)
                {
                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.Id,
                        id);
                }

                writer.WriteEndElement(); // Ids
            }
        }

        /// <summary>
        /// Gets or sets the ids of the time zones that should be returned by the server.
        /// </summary>
        internal IEnumerable<string> Ids
        {
            get { return this.ids; }
            set { this.ids = value; }
        }
    }
}