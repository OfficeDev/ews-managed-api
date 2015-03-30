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
    /// <summary>
    /// Represents a GetInboxRules request.
    /// </summary>
    internal sealed class GetInboxRulesRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// The smtp address of the mailbox from which to get the inbox rules.
        /// </summary>
        private string mailboxSmtpAddress;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetInboxRulesRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetInboxRulesRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the address of the mailbox from which to get the inbox rules.
        /// </summary>
        internal string MailboxSmtpAddress 
        {
            get { return this.mailboxSmtpAddress; }
            set { this.mailboxSmtpAddress = value; }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetInboxRules;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.mailboxSmtpAddress))
            {
                writer.WriteElementValue(
                    XmlNamespace.Messages, 
                    XmlElementNames.MailboxSmtpAddress, 
                    this.mailboxSmtpAddress);
            }
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetInboxRulesResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetInboxRulesResponse response = new GetInboxRulesResponse();
            response.LoadFromXml(reader, XmlElementNames.GetInboxRulesResponse);
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010_SP1;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal GetInboxRulesResponse Execute()
        {
            GetInboxRulesResponse serviceResponse = (GetInboxRulesResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}