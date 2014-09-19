// ---------------------------------------------------------------------------
// <copyright file="GetUserConfigurationResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUserConfigurationResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Text;

    /// <summary>
    /// Represents a response to a GetUserConfiguration request.
    /// </summary>
    internal sealed class GetUserConfigurationResponse : ServiceResponse
    {
        private UserConfiguration userConfiguration;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserConfigurationResponse"/> class.
        /// </summary>
        /// <param name="userConfiguration">The userConfiguration.</param>
        internal GetUserConfigurationResponse(UserConfiguration userConfiguration)
            : base()
        {
            EwsUtilities.Assert(
                userConfiguration != null,
                "GetUserConfigurationResponse.ctor",
                "userConfiguration is null");

            this.userConfiguration = userConfiguration;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.userConfiguration.LoadFromXml(reader);
        }

        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            this.UserConfiguration.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.UserConfiguration), service);
        }

        /// <summary>
        /// Gets the user configuration that was created.
        /// </summary>
        public UserConfiguration UserConfiguration
        {
            get { return this.userConfiguration; }
        }
    }
}
