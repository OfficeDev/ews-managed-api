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
    using System.Collections.Generic;

    /// <summary>
    /// Represents a RemoveDelete request.
    /// </summary>
    internal class RemoveDelegateRequest : DelegateManagementRequestBase<DelegateManagementResponse>
    {
        private List<UserId> userIds = new List<UserId>();

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoveDelegateRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal RemoveDelegateRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Asserts the valid.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParamCollection(this.UserIds, "UserIds");
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            base.WriteElementsToXml(writer);

            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.UserIds);

            foreach (UserId userId in this.UserIds)
            {
                userId.WriteToXml(writer, XmlElementNames.UserId);
            }

            writer.WriteEndElement(); // UserIds
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.RemoveDelegateResponse;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.RemoveDelegate;
        }

        /// <summary>
        /// Creates the response.
        /// </summary>
        /// <returns>Service response.</returns>
        internal override DelegateManagementResponse CreateResponse()
        {
            return new DelegateManagementResponse(false, null);
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
        /// Gets the user ids.
        /// </summary>
        /// <value>The user ids.</value>
        public List<UserId> UserIds
        {
            get { return this.userIds; }
        }
    }
}