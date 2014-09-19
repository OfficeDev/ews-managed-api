// ---------------------------------------------------------------------------
// <copyright file="UpdateUserConfigurationRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UpdateUserConfigurationRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a UpdateUserConfiguration request.
    /// </summary>
    internal class UpdateUserConfigurationRequest : MultiResponseServiceRequest<ServiceResponse>, IJsonSerializable
    {
        protected UserConfiguration userConfiguration;

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.userConfiguration, "userConfiguration");
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ServiceResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new ServiceResponse();
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2010;
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
            return XmlElementNames.UpdateUserConfiguration;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.UpdateUserConfigurationResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.UpdateUserConfigurationResponseMessage;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            // Write UserConfiguation element
            this.userConfiguration.WriteToXml(writer, XmlNamespace.Messages, XmlElementNames.UserConfiguration);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateUserConfigurationRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal UpdateUserConfigurationRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
        {
        }

        /// <summary>
        /// Gets or sets the user configuration.
        /// </summary>
        /// <value>The user configuration.</value>
        public UserConfiguration UserConfiguration
        {
            get { return this.userConfiguration; }
            set { this.userConfiguration = value; }
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonObject = new JsonObject();

            jsonObject.Add(XmlElementNames.UserConfiguration, ((IJsonSerializable)this.UserConfiguration).ToJson(service));

            return jsonObject;
        }
    }
}
