// ---------------------------------------------------------------------------
// <copyright file="GetUnifiedGroupUnseenCountResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetUnifiedGroupUnseenCountResponse class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Represents a response to the GetUnifiedGroupUnseenCount operation
    /// </summary>
    internal sealed class GetUnifiedGroupUnseenCountResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetUnifiedGroupUnseenCountResponse"/> class.
        /// </summary>
        internal GetUnifiedGroupUnseenCountResponse() :
             base()
        {
        }

        /// <summary>
        /// Gets or sets the unseen count
        /// </summary>
        public int UnseenCount { get; set; }

        /// <summary>
        /// Read response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);
            this.UnseenCount = reader.ReadElementValue<int>(XmlNamespace.NotSpecified, XmlElementNames.UnseenCount);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            if (responseObject.ContainsKey(XmlElementNames.UnseenCount))
            {
                this.UnseenCount = responseObject.ReadAsInt(XmlElementNames.UnseenCount);
            }
        }
    }
}
