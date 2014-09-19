// ---------------------------------------------------------------------------
// <copyright file="SetUserOofSettingsRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetUserOofSettingsRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a SetUserOofSettings request.
    /// </summary>
    internal sealed class SetUserOofSettingsRequest : SimpleServiceRequestBase
    {
        private string smtpAddress;
        private OofSettings oofSettings;

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SetUserOofSettingsRequest;
        }

        /// <summary>
        /// Validate request..
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            EwsUtilities.ValidateParam(this.SmtpAddress, "SmtpAddress");
            EwsUtilities.ValidateParam(this.OofSettings, "OofSettings");
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Mailbox);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Address, this.SmtpAddress);
            writer.WriteEndElement(); // Mailbox

            this.OofSettings.WriteToXml(writer, XmlElementNames.UserOofSettings);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SetUserOofSettingsResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Service response.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            ServiceResponse serviceResponse = new ServiceResponse();

            serviceResponse.LoadFromXml(reader, XmlElementNames.ResponseMessage);

            return serviceResponse;
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
        /// Initializes a new instance of the <see cref="SetUserOofSettingsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SetUserOofSettingsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal ServiceResponse Execute()
        {
            ServiceResponse serviceResponse = (ServiceResponse)this.InternalExecute();

            serviceResponse.ThrowIfNecessary();

            return serviceResponse;
        }

        /// <summary>
        /// Gets or sets the SMTP address.
        /// </summary>
        public string SmtpAddress
        {
            get { return this.smtpAddress; }
            set { this.smtpAddress = value; }
        }

        /// <summary>
        /// Gets or sets the oof settings.
        /// </summary>
        public OofSettings OofSettings
        {
            get { return this.oofSettings; }
            set { this.oofSettings = value; }
        }
    }
}
