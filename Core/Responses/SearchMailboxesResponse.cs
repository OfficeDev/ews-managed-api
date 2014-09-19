// ---------------------------------------------------------------------------
// <copyright file="SearchMailboxesResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchMailboxesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the SearchMailboxes response.
    /// </summary>
    public sealed class SearchMailboxesResponse : ServiceResponse
    {
        SearchMailboxesResult searchResult = null;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchMailboxesResponse"/> class.
        /// </summary>
        internal SearchMailboxesResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.searchResult = new SearchMailboxesResult();

            base.ReadElementsFromXml(reader);

            this.searchResult = SearchMailboxesResult.LoadFromXml(reader);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            if (responseObject.ContainsKey(XmlElementNames.SearchMailboxesResult))
            {
                JsonObject jsonSearchResult = responseObject.ReadAsJsonObject(XmlElementNames.SearchMailboxesResult);

                this.searchResult = SearchMailboxesResult.LoadFromJson(jsonSearchResult);
            }
        }

        /// <summary>
        /// Search mailboxes result
        /// </summary>
        public SearchMailboxesResult SearchResult
        {
            get { return this.searchResult; }
            internal set { this.searchResult = value; }
        }
    }
}
