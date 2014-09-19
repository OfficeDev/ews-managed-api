// ---------------------------------------------------------------------------
// <copyright file="GetDiscoverySearchConfigurationResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetDiscoverySearchConfigurationResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the GetDiscoverySearchConfiguration response.
    /// </summary>
    public sealed class GetDiscoverySearchConfigurationResponse : ServiceResponse
    {
        List<DiscoverySearchConfiguration> configurations = new List<DiscoverySearchConfiguration>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetDiscoverySearchConfigurationResponse"/> class.
        /// </summary>
        internal GetDiscoverySearchConfigurationResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.configurations.Clear();

            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.DiscoverySearchConfigurations);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();
                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.DiscoverySearchConfiguration))
                    {
                        this.configurations.Add(DiscoverySearchConfiguration.LoadFromXml(reader));
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.DiscoverySearchConfigurations));
            }
            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.DiscoverySearchConfigurations);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            this.configurations.Clear();

            base.ReadElementsFromJson(responseObject, service);

            if (responseObject.ContainsKey(XmlElementNames.DiscoverySearchConfigurations))
            {
                foreach (object searchConfiguration in responseObject.ReadAsArray(XmlElementNames.DiscoverySearchConfigurations))
                {
                    JsonObject jsonSearchConfiguration = searchConfiguration as JsonObject;
                    this.configurations.Add(DiscoverySearchConfiguration.LoadFromJson(jsonSearchConfiguration));
                }
            }
        }

        /// <summary>
        /// Searchable mailboxes result
        /// </summary>
        public DiscoverySearchConfiguration[] DiscoverySearchConfigurations
        {
            get { return this.configurations.ToArray(); }
        }
    }
}
