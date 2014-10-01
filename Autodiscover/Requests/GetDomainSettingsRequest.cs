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
// <summary>Defines the GetDomainSettingsRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a GetDomainSettings request.
    /// </summary>
    internal class GetDomainSettingsRequest : AutodiscoverRequest
    {
        /// <summary>
        /// Action Uri of Autodiscover.GetDomainSettings method.
        /// </summary>
        private const string GetDomainSettingsActionUri = EwsUtilities.AutodiscoverSoapNamespace + "/Autodiscover/GetDomainSettings";

        private List<string> domains;
        private List<DomainSettingName> settings;
        private ExchangeVersion? requestedVersion;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetDomainSettingsRequest"/> class.
        /// </summary>
        /// <param name="service">Autodiscover service associated with this request.</param>
        /// <param name="url">URL of Autodiscover service.</param>
        internal GetDomainSettingsRequest(AutodiscoverService service, Uri url)
            : base(service, url)
        {
        }

        /// <summary>
        /// Validates the request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            EwsUtilities.ValidateParam(this.Domains, "domains");
            EwsUtilities.ValidateParam(this.Settings, "settings");

            if (this.Settings.Count == 0)
            {
                throw new ServiceValidationException(Strings.InvalidAutodiscoverSettingsCount);
            }

            if (domains.Count == 0)
            {
                throw new ServiceValidationException(Strings.InvalidAutodiscoverDomainsCount);
            }

            foreach (string domain in this.domains)
            {
                if (string.IsNullOrEmpty(domain))
                {
                    throw new ServiceValidationException(Strings.InvalidAutodiscoverDomain);
                }
            }
        }

        /// <summary>
        /// Executes this instance.
        /// </summary>
        /// <returns></returns>
        internal GetDomainSettingsResponseCollection Execute()
        {
            GetDomainSettingsResponseCollection responses = (GetDomainSettingsResponseCollection)this.InternalExecute();
            if (responses.ErrorCode == AutodiscoverErrorCode.NoError)
            {
                this.PostProcessResponses(responses);
            }
            return responses;
        }

        /// <summary>
        /// Post-process responses to GetDomainSettings.
        /// </summary>
        /// <param name="responses">The GetDomainSettings responses.</param>
        private void PostProcessResponses(GetDomainSettingsResponseCollection responses)
        {
            // Note:The response collection may not include all of the requested domains if the request has been throttled.
            for (int index = 0; index < responses.Count; index++)
            {
                responses[index].Domain = this.Domains[index];
            }
        }

        /// <summary>
        /// Gets the name of the request XML element.
        /// </summary>
        /// <returns>Request XML element name.</returns>
        internal override string GetRequestXmlElementName()
        {
            return XmlElementNames.GetDomainSettingsRequestMessage;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>Response XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetDomainSettingsResponseMessage;
        }

        /// <summary>
        /// Gets the WS-Addressing action name.
        /// </summary>
        /// <returns>WS-Addressing action name.</returns>
        internal override string GetWsAddressingActionName()
        {
            return GetDomainSettingsActionUri;
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <returns>AutodiscoverResponse</returns>
        internal override AutodiscoverResponse CreateServiceResponse()
        {
            return new GetDomainSettingsResponseCollection();
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(
                "xmlns",
                EwsUtilities.AutodiscoverSoapNamespacePrefix,
                EwsUtilities.AutodiscoverSoapNamespace);
        }

        /// <summary>
        /// Writes request to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.Request);

            writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.Domains);

            foreach (string domain in this.Domains)
            {
                if (!string.IsNullOrEmpty(domain))
                {
                    writer.WriteElementValue(
                        XmlNamespace.Autodiscover,
                        XmlElementNames.Domain,
                        domain);
                }
            }
            writer.WriteEndElement(); // Domains

            writer.WriteStartElement(XmlNamespace.Autodiscover, XmlElementNames.RequestedSettings);
            foreach (DomainSettingName setting in settings)
            {
                writer.WriteElementValue(
                    XmlNamespace.Autodiscover,
                    XmlElementNames.Setting,
                    setting);
            }

            writer.WriteEndElement(); // RequestedSettings

            if (this.requestedVersion.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Autodiscover, XmlElementNames.RequestedVersion, this.requestedVersion.Value);
            }

            writer.WriteEndElement(); // Request
        }

        /// <summary>
        /// Gets or sets the domains.
        /// </summary>
        internal List<string> Domains
        {
            get { return domains; }
            set { domains = value; }
        }

        /// <summary>
        /// Gets or sets the settings.
        /// </summary>
        internal List<DomainSettingName> Settings
        {
            get { return settings; }
            set { settings = value; }
        }

        /// <summary>
        /// Gets or sets the RequestedVersion.
        /// </summary>
        internal ExchangeVersion? RequestedVersion
        {
            get { return requestedVersion; }
            set { requestedVersion = value; }
        }
    }
}
