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
    using System.Net;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a request of a get user photo operation
    /// </summary>
    internal sealed class SetUserPhotoRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service">Exchange web service</param>
        internal SetUserPhotoRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// email address accessor
        /// </summary>
        internal string EmailAddress { get; set; }

        internal byte[] Photo { get; set; }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            if (string.IsNullOrEmpty(this.EmailAddress))
            {
                throw new ServiceLocalException(Strings.InvalidEmailAddress);
            }

            if (this.Photo == null || this.Photo.Length <= 0)
            {
                throw new ServiceLocalException(Strings.UserPhotoNotSpecified);
            }

            base.Validate();
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            // Emit the EmailAddress element
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Email);
            writer.WriteValue(this.EmailAddress, XmlElementNames.Email);
            writer.WriteEndElement();

            string encodedPhoto = Convert.ToBase64String(this.Photo);

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Content);
            writer.WriteValue(encodedPhoto, XmlElementNames.Content);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Adds header values to the request
        /// </summary>
        /// <param name="webHeaderCollection">The collection of headers to add to</param>
        internal override void AddHeaders(WebHeaderCollection webHeaderCollection)
        {
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="responseHeaders">The HTTP response headers</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader, WebHeaderCollection responseHeaders)
        {
            SetUserPhotoResponse response = new SetUserPhotoResponse();
            response.LoadFromXml(reader, XmlElementNames.SetUserPhotoResponse);
            response.ReadHeader(responseHeaders);
            return response;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SetUserPhoto;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SetUserPhotoResponse;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2016;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal SetUserPhotoResponse Execute()
        {
            return SetUserPhotoRequest.SetResultOrDefault(this.InternalExecute);
        }

                /// <summary>
        /// Ends executing this async request.
        /// </summary>
        /// <param name="asyncResult">The async result</param>
        /// <returns>Service response collection.</returns>
        internal SetUserPhotoResponse EndExecute(IAsyncResult asyncResult)
        {
            return SetUserPhotoRequest.SetResultOrDefault(() => this.EndInternalExecute(asyncResult));
        }

        private static SetUserPhotoResponse SetResultOrDefault(Func<object> serviceResponseFactory)
        {
            try
            {
                return (SetUserPhotoResponse)serviceResponseFactory();
            }
            catch (ServiceRequestException ex)
            {
                throw;
            }
        }
    }
}