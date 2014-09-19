// ---------------------------------------------------------------------------
// <copyright file="GetNonIndexableItemDetailsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetNonIndexableItemDetailsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the GetNonIndexableItemDetails response.
    /// </summary>
    public sealed class GetNonIndexableItemDetailsResponse : ServiceResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetNonIndexableItemDetailsResponse"/> class.
        /// </summary>
        internal GetNonIndexableItemDetailsResponse()
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

            this.NonIndexableItemsResult = NonIndexableItemDetailsResult.LoadFromXml(reader);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            throw new NotImplementedException("GetNonIndexableItemdDetails doesn't support JSON.");
        }

        /// <summary>
        /// Non indexable item result
        /// </summary>
        public NonIndexableItemDetailsResult NonIndexableItemsResult
        {
            get;
            internal set;
        }
    }
}
