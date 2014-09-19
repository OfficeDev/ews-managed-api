// ---------------------------------------------------------------------------
// <copyright file="GetEventsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetEventsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a subscription event retrieval operation.
    /// </summary>
    internal sealed class GetEventsResponse : ServiceResponse
    {
        private GetEventsResults results = new GetEventsResults();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetEventsResponse"/> class.
        /// </summary>
        internal GetEventsResponse()
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

            this.results.LoadFromXml(reader);
        }

        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            this.results.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.Notification), service);
        }

        /// <summary>
        /// Gets event results from subscription.
        /// </summary>
        internal GetEventsResults Results
        {
            get { return this.results; }
        }
    }
}
