//-----------------------------------------------------------------------
// <copyright file="MarkAsJunkResponse.cs" company="Microsoft Corp.">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Definition for MarkAsJunkResponse
    /// </summary>
    public class MarkAsJunkResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetItemResponse"/> class.
        /// </summary>
        internal MarkAsJunkResponse() : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.MovedItemId))
            {
                this.MovedItemId = new ItemId();
                this.MovedItemId.LoadFromXml(reader, XmlNamespace.Messages, XmlElementNames.MovedItemId);

                reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.MovedItemId);
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">Json response object</param>
        /// <param name="service">Exchange service</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);
            if (responseObject.ContainsKey(XmlElementNames.Token))
            {
                this.MovedItemId = new ItemId();
                this.MovedItemId.LoadFromJson(responseObject.ReadAsJsonObject(XmlElementNames.MovedItemId), service);
            }
        }

        /// <summary>
        /// Gets the moved item id.
        /// </summary>
        public ItemId MovedItemId { get; private set; }
    }
}
