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

    /// <summary>
    /// Represents a request of a find persona operation
    /// </summary>
    internal sealed class FindPeopleRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="service">Exchange web service</param>
        internal FindPeopleRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Accessors of the view controlling the number of personas returned.
        /// </summary>
        internal ViewBase View { get; set; }

        /// <summary>
        /// Folder Id accessors
        /// </summary>
        internal FolderId FolderId { get; set; }

        /// <summary>
        /// Search filter accessors
        /// Available search filter classes include SearchFilter.IsEqualTo,
        /// SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection. If SearchFilter
        /// is null, no search filters are applied.
        /// </summary>
        internal SearchFilter SearchFilter { get; set; }

        /// <summary>
        /// Query string accessors
        /// </summary>
        internal string QueryString { get; set; }

        /// <summary>
        /// Whether to search the people suggestion index
        /// </summary>
        internal bool SearchPeopleSuggestionIndex { get; set; }

        /// <summary>
        /// The context for suggestion index enabled queries
        /// </summary>
        internal Dictionary<string, string> Context { get; set; }

        /// <summary>
        /// The query mode for suggestion index enabled queries
        /// </summary>
        internal PeopleQueryMode QueryMode { get; set; }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            this.View.InternalValidate(this);
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
            this.View.WriteAttributesToXml(writer);
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.SearchFilter != null)
            {
                // Emit the Restriction element
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.Restriction);
                this.SearchFilter.WriteToXml(writer);
                writer.WriteEndElement();
            }

            // Emit the View element
            this.View.WriteToXml(writer, null);

            // Emit the SortOrder
            this.View.WriteOrderByToXml(writer);

            // Emit the ParentFolderId element
            if (this.FolderId != null)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.ParentFolderId);
                this.FolderId.WriteToXml(writer);
                writer.WriteEndElement();
            }

            if (!string.IsNullOrEmpty(this.QueryString))
            {
                // Emit the QueryString element
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.QueryString);
                writer.WriteValue(this.QueryString, XmlElementNames.QueryString);
                writer.WriteEndElement();
            }

            // Emit the SuggestionIndex-enabled elements
            if (this.SearchPeopleSuggestionIndex)
            {
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SearchPeopleSuggestionIndex);
                writer.WriteValue(this.SearchPeopleSuggestionIndex.ToString().ToLowerInvariant(), XmlElementNames.SearchPeopleSuggestionIndex);
                writer.WriteEndElement();

                // Write the Context key value pairs
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SearchPeopleContext);
                foreach (KeyValuePair<string, string> contextItem in this.Context)
                {
                    writer.WriteStartElement(XmlNamespace.Types, "ContextProperty");

                    writer.WriteStartElement(XmlNamespace.Types, "Key");
                    writer.WriteValue(contextItem.Key, "Key");
                    writer.WriteEndElement();

                    writer.WriteStartElement(XmlNamespace.Types, "Value");
                    writer.WriteValue(contextItem.Value, "Value");
                    writer.WriteEndElement();

                    writer.WriteEndElement();
                }
                writer.WriteEndElement();

                // Write the Query Mode Sources
                writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.SearchPeopleQuerySources);
                foreach (string querySource in this.QueryMode.Sources)
                {
                    writer.WriteStartElement(XmlNamespace.Types, "Source");
                    writer.WriteValue(querySource, "Source");
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            if (this.Service.RequestedServerVersion >= this.GetMinimumRequiredServerVersion())
            {
                if (this.View.PropertySet != null)
                {
                    this.View.PropertySet.WriteToXml(writer, ServiceObjectType.Persona);
                }
            }
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            FindPeopleResponse response = new FindPeopleResponse();
            response.LoadFromXml(reader, XmlElementNames.FindPeopleResponse);
            return response;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.FindPeople;
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.FindPeopleResponse;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013_SP1;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal FindPeopleResponse Execute()
        {
            FindPeopleResponse serviceResponse = (FindPeopleResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}