#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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
