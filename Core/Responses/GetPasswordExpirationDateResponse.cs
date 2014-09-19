// ---------------------------------------------------------------------------
// <copyright file="GetPasswordExpirationDateResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetPasswordExpirationDateResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a GetPasswordExpirationDate operation
    /// </summary>
    internal sealed class GetPasswordExpirationDateResponse : ServiceResponse
    {
        private DateTime? passwordExpirationDate;    

        /// <summary>
        /// Initializes a new instance of the <see cref="GetPasswordExpirationDateResponse"/> class.
        /// </summary>
        internal GetPasswordExpirationDateResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);
            this.passwordExpirationDate = reader.ReadElementValueAsDateTime(XmlNamespace.NotSpecified, XmlElementNames.PasswordExpirationDate);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);
            this.passwordExpirationDate = service.ConvertUniversalDateTimeStringToLocalDateTime(responseObject.ReadAsString(XmlElementNames.PasswordExpirationDate)).Value;
        }

        /// <summary>
        /// Password expiration date
        /// </summary>
        public DateTime? PasswordExpirationDate
        {
            get { return this.passwordExpirationDate; }
        }
    }
}
