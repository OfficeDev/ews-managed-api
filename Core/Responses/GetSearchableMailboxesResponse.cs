// ---------------------------------------------------------------------------
// <copyright file="GetSearchableMailboxesResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetSearchableMailboxesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the GetSearchableMailboxes response.
    /// </summary>
    public sealed class GetSearchableMailboxesResponse : ServiceResponse
    {
        List<SearchableMailbox> searchableMailboxes = new List<SearchableMailbox>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetSearchableMailboxesResponse"/> class.
        /// </summary>
        internal GetSearchableMailboxesResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.searchableMailboxes.Clear();

            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.SearchableMailboxes);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();
                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.SearchableMailbox))
                    {
                        this.searchableMailboxes.Add(SearchableMailbox.LoadFromXml(reader));
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.SearchableMailboxes));
            }

            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.FailedMailboxes))
            {
                this.FailedMailboxes = FailedSearchMailbox.LoadFailedMailboxesXml(XmlNamespace.Messages, reader);
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            this.searchableMailboxes.Clear();

            base.ReadElementsFromJson(responseObject, service);

            if (responseObject.ContainsKey(XmlElementNames.SearchMailboxes))
            {
                foreach (object searchableMailboxObject in responseObject.ReadAsArray(XmlElementNames.SearchableMailboxes))
                {
                    JsonObject jsonSearchableMailbox = searchableMailboxObject as JsonObject;
                    this.searchableMailboxes.Add(SearchableMailbox.LoadFromJson(jsonSearchableMailbox));
                }
            }
        }

        /// <summary>
        /// Searchable mailboxes result
        /// </summary>
        public SearchableMailbox[] SearchableMailboxes
        {
            get { return this.searchableMailboxes.ToArray(); }
        }

        /// <summary>
        /// Failed mailboxes
        /// </summary>
        public FailedSearchMailbox[] FailedMailboxes { get; set; }
    }
}
