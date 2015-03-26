// ---------------------------------------------------------------------------
// <copyright file="SetOMEConfigurationRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetOMEConfigurationRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a SetOMEConfiguration request.
    /// </summary>
    internal sealed class SetOMEConfigurationRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// The XML representation of EncryptionConfigurationData
        /// </summary>
        private readonly string xml;

        /// <summary>
        /// The XML representation of EncryptionConfigurationData
        /// </summary>
        public string Xml
        {
            get { return this.xml; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SetOMEConfigurationRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xml">The XML representation of EncryptionConfigurationData</param>
        internal SetOMEConfigurationRequest(ExchangeService service, string xml) : base(service)
        {
            this.xml = xml;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SetOMEConfigurationRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.OMEConfigurationXml, this.Xml);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SetOMEConfigurationResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            SetOMEConfigurationResponse response = new SetOMEConfigurationResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal ServiceResponse Execute()
        {
            SetOMEConfigurationResponse serviceResponse = (SetOMEConfigurationResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}