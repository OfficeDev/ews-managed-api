/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a SearchMailboxesRequest request.
    /// </summary>
    internal sealed class SearchMailboxesRequest : MultiResponseServiceRequest<SearchMailboxesResponse>, IDiscoveryVersionable
    {
        private List<MailboxQuery> searchQueries = new List<MailboxQuery>();
        private SearchResultType searchResultType = SearchResultType.PreviewOnly;
        private SortDirection sortOrder = SortDirection.Ascending;
        private string sortByProperty;
        private bool performDeduplication;
        private int pageSize;
        private string pageItemReference;
        private SearchPageDirection pageDirection = SearchPageDirection.Next;
        private PreviewItemResponseShape previewItemResponseShape;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchMailboxesRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal SearchMailboxesRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override SearchMailboxesResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new SearchMailboxesResponse();
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SearchMailboxesResponse;
        }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseMessageXmlElementName()
        {
            return XmlElementNames.SearchMailboxesResponseMessage;
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return 1;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SearchMailboxes;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.SearchQueries == null || this.SearchQueries.Count == 0)
            {
                throw new ServiceValidationException(Strings.MailboxQueriesParameterIsNotSpecified);
            }

            foreach (MailboxQuery searchQuery in this.SearchQueries)
            {
                if (searchQuery.MailboxSearchScopes == null || searchQuery.MailboxSearchScopes.Length == 0)
                {
                    throw new ServiceValidationException(Strings.MailboxQueriesParameterIsNotSpecified);
                }

                foreach (MailboxSearchScope searchScope in searchQuery.MailboxSearchScopes)
                {
                    if (searchScope.ExtendedAttributes != null && searchScope.ExtendedAttributes.Count > 0 && !DiscoverySchemaChanges.SearchMailboxesExtendedData.IsCompatible(this))
                    {
                        throw new ServiceVersionException(
                           string.Format(
                                         Strings.ClassIncompatibleWithRequestVersion,
                                         typeof(ExtendedAttribute).Name,
                                         DiscoverySchemaChanges.SearchMailboxesExtendedData.MinimumServerVersion));
                    }

                    if (searchScope.SearchScopeType != MailboxSearchScopeType.LegacyExchangeDN && (!DiscoverySchemaChanges.SearchMailboxesExtendedData.IsCompatible(this) || !DiscoverySchemaChanges.SearchMailboxesAdditionalSearchScopes.IsCompatible(this)))
                    {
                        throw new ServiceVersionException(
                           string.Format(
                                         Strings.EnumValueIncompatibleWithRequestVersion,
                                         searchScope.SearchScopeType.ToString(),
                                         typeof(MailboxSearchScopeType).Name,
                                         DiscoverySchemaChanges.SearchMailboxesAdditionalSearchScopes.MinimumServerVersion));
                    }
                }
            }

            if (!string.IsNullOrEmpty(this.SortByProperty))
            {
                PropertyDefinitionBase prop = null;
                try
                {
                    prop = ServiceObjectSchema.FindPropertyDefinition(this.SortByProperty);
                }
                catch (KeyNotFoundException)
                {
                }

                if (prop == null)
                {
                    throw new ServiceValidationException(string.Format(Strings.InvalidSortByPropertyForMailboxSearch, this.SortByProperty));
                }
            }
        }

        /// <summary>
        /// Parses the response.
        /// See O15:324151 on why we need to override ParseResponse here instead of calling the one in MultiResponseServiceRequest.cs
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Service response collection.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            ServiceResponseCollection<SearchMailboxesResponse> serviceResponses = new ServiceResponseCollection<SearchMailboxesResponse>();

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

            while (true)
            {
                // Read ahead to see if we've reached the end of the response messages early.
                reader.Read();
                if (reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages))
                {
                    break;
                }

                SearchMailboxesResponse response = new SearchMailboxesResponse();
                response.LoadFromXml(reader, this.GetResponseMessageXmlElementName());
                serviceResponses.Add(response);
            }

            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

            return serviceResponses;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SearchQueries);
            foreach (MailboxQuery mailboxQuery in this.SearchQueries)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MailboxQuery);
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Query, mailboxQuery.Query);
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MailboxSearchScopes);
                foreach (MailboxSearchScope mailboxSearchScope in mailboxQuery.MailboxSearchScopes)
                {
                    // The checks here silently downgrade the schema based on compatibility checks, to receive errors use the validate method
                    if (mailboxSearchScope.SearchScopeType == MailboxSearchScopeType.LegacyExchangeDN || DiscoverySchemaChanges.SearchMailboxesAdditionalSearchScopes.IsCompatible(this))
                    {
                        writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.MailboxSearchScope);
                        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Mailbox, mailboxSearchScope.Mailbox);
                        writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SearchScope, mailboxSearchScope.SearchScope);

                        if (DiscoverySchemaChanges.SearchMailboxesExtendedData.IsCompatible(this))
                        {
                            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ExtendedAttributes);

                            if (mailboxSearchScope.SearchScopeType != MailboxSearchScopeType.LegacyExchangeDN)
                            {
                                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ExtendedAttribute);
                                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ExtendedAttributeName, XmlElementNames.SearchScopeType);
                                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ExtendedAttributeValue, mailboxSearchScope.SearchScopeType);
                                writer.WriteEndElement();
                            }

                            if (mailboxSearchScope.ExtendedAttributes != null && mailboxSearchScope.ExtendedAttributes.Count > 0)
                            {
                                foreach (ExtendedAttribute attribute in mailboxSearchScope.ExtendedAttributes)
                                {
                                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ExtendedAttribute);
                                    writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ExtendedAttributeName, attribute.Name);
                                    writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.ExtendedAttributeValue, attribute.Value);
                                    writer.WriteEndElement();
                                }
                            }

                            writer.WriteEndElement();  // ExtendedData
                        }

                        writer.WriteEndElement();   // MailboxSearchScope
                    }
                }

                writer.WriteEndElement();   // MailboxSearchScopes
                writer.WriteEndElement();   // MailboxQuery
            }

            writer.WriteEndElement();   // SearchQueries
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.ResultType, this.ResultType);

            if (this.PreviewItemResponseShape != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.PreviewItemResponseShape);
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.BaseShape, this.PreviewItemResponseShape.BaseShape);
                if (this.PreviewItemResponseShape.AdditionalProperties != null && this.PreviewItemResponseShape.AdditionalProperties.Length > 0)
                {
                    writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.AdditionalProperties);
                    foreach (ExtendedPropertyDefinition additionalProperty in this.PreviewItemResponseShape.AdditionalProperties)
                    {
                        additionalProperty.WriteToXml(writer);
                    }

                    writer.WriteEndElement();   // AdditionalProperties
                }

                writer.WriteEndElement();   // PreviewItemResponseShape
            }

            if (!string.IsNullOrEmpty(this.SortByProperty))
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SortBy);
                writer.WriteAttributeValue(XmlElementNames.Order, this.SortOrder.ToString());
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.FieldURI);
                writer.WriteAttributeValue(XmlElementNames.FieldURI, this.sortByProperty);
                writer.WriteEndElement();   // FieldURI
                writer.WriteEndElement();   // SortBy
            }

            // Language
            if (!string.IsNullOrEmpty(this.Language))
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.Language, this.Language);
            }

            // Dedupe
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.Deduplication, this.performDeduplication);

            if (this.PageSize > 0)
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.PageSize, this.PageSize.ToString());
            }

            if (!string.IsNullOrEmpty(this.PageItemReference))
            {
                writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.PageItemReference, this.PageItemReference);
            }

            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.PageDirection, this.PageDirection.ToString());
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Collection of query + mailboxes
        /// </summary>
        public List<MailboxQuery> SearchQueries
        {
            get { return this.searchQueries; }
            set { this.searchQueries = value; }
        }

        /// <summary>
        /// Search result type
        /// </summary>
        public SearchResultType ResultType
        {
            get { return this.searchResultType; }
            set { this.searchResultType = value; }
        }

        /// <summary>
        /// Preview item response shape
        /// </summary>
        public PreviewItemResponseShape PreviewItemResponseShape
        {
            get { return this.previewItemResponseShape; }
            set { this.previewItemResponseShape = value; }
        }

        /// <summary>
        /// Sort order
        /// </summary>
        public SortDirection SortOrder
        {
            get { return this.sortOrder; }
            set { this.sortOrder = value; }
        }

        /// <summary>
        /// Sort by property name
        /// </summary>
        public string SortByProperty
        {
            get { return this.sortByProperty; }
            set { this.sortByProperty = value; }
        }

        /// <summary>
        /// Query language
        /// </summary>
        public string Language
        {
            get;
            set;
        }

        /// <summary>
        /// Perform deduplication or not
        /// </summary>
        public bool PerformDeduplication
        {
            get { return this.performDeduplication; }
            set { this.performDeduplication = value; }
        }

        /// <summary>
        /// Page size
        /// </summary>
        public int PageSize
        {
            get { return this.pageSize; }
            set { this.pageSize = value; }
        }

        /// <summary>
        /// Page item reference
        /// </summary>
        public string PageItemReference
        {
            get { return this.pageItemReference; }
            set { this.pageItemReference = value; }
        }

        /// <summary>
        /// Page direction
        /// </summary>
        public SearchPageDirection PageDirection
        {
            get { return this.pageDirection; }
            set { this.pageDirection = value; }
        }

        /// <summary>
        /// Gets or sets the server version.
        /// </summary>
        /// <value>
        /// The server version.
        /// </value>
        long IDiscoveryVersionable.ServerVersion
        {
            get;
            set;
        }
    }
}