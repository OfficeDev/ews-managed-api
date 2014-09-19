// ---------------------------------------------------------------------------
// <copyright file="GetInboxRulesRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetInboxRulesRequest class.</summary>
//-----------------------------------------------------------------------

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