// ---------------------------------------------------------------------------
// <copyright file="RegisterConsentRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;

    using Microsoft.Exchange.WebServices.Data.Enumerations;

    /// <summary>
    /// Represents a RegisterConsent request.
    /// </summary>
    internal sealed class RegisterConsentRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RegisterConsentRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="id">Extension id.</param>
        /// <param name="state">Sets the consent state of an extension.</param>
        internal RegisterConsentRequest(ExchangeService service, string id, ConsentState state)
            : base(service)
        {
            this.Id = id;
            this.ConsentState = state;
        }

        /// <summary>
        /// Extension id
        /// </summary>
        private string Id
        {
            get;
            set;
        }

        /// <summary>
        /// User decision on the consent state of an extension
        /// </summary>
        private ConsentState ConsentState
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.DisableAppRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ID, this.Id);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.AcceptanceState, this.ConsentState);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.RegisterConsentResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            RegisterConsentResponse response = new RegisterConsentResponse();
            response.LoadFromXml(reader, XmlElementNames.RegisterConsentResponse);
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
        internal RegisterConsentResponse Execute()
        {
            RegisterConsentResponse serviceResponse = (RegisterConsentResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}