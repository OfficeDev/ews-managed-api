// ---------------------------------------------------------------------------
// <copyright file="SearchFilter.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchFilter class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the base search filter class. Use descendant search filter classes such as SearchFilter.IsEqualTo,
    /// SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection to define search filters.
    /// </summary>
    public abstract partial class SearchFilter : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SearchFilter"/> class.
        /// </summary>
        internal SearchFilter()
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>SearchFilter.</returns>
        internal static SearchFilter LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.EnsureCurrentNodeIsStartElement();

            string localName = reader.LocalName;

            SearchFilter searchFilter = GetSearchFilterInstance(localName);

            if (searchFilter != null)
            {
                searchFilter.LoadFromXml(reader, reader.LocalName);
            }

            return searchFilter;
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        /// <returns></returns>
        internal static SearchFilter LoadSearchFilterFromJson(JsonObject jsonObject, ExchangeService service)
        {
            SearchFilter searchFilter = GetSearchFilterInstance(jsonObject.ReadTypeString());

            if (searchFilter != null)
            {
                searchFilter.LoadFromJson(jsonObject, service);
            }

            return searchFilter;
        }

        /// <summary>
        /// Gets the search filter instance.
        /// </summary>
        /// <param name="localName">Name of the local.</param>
        /// <returns></returns>
        private static SearchFilter GetSearchFilterInstance(string localName)
        {
            SearchFilter searchFilter;
            switch (localName)
            {
                case XmlElementNames.Exists:
                    searchFilter = new Exists();
                    break;
                case XmlElementNames.Contains:
                    searchFilter = new ContainsSubstring();
                    break;
                case XmlElementNames.Excludes:
                    searchFilter = new ExcludesBitmask();
                    break;
                case XmlElementNames.Not:
                    searchFilter = new Not();
                    break;
                case XmlElementNames.And:
                    searchFilter = new SearchFilterCollection(LogicalOperator.And);
                    break;
                case XmlElementNames.Or:
                    searchFilter = new SearchFilterCollection(LogicalOperator.Or);
                    break;
                case XmlElementNames.IsEqualTo:
                    searchFilter = new IsEqualTo();
                    break;
                case XmlElementNames.IsNotEqualTo:
                    searchFilter = new IsNotEqualTo();
                    break;
                case XmlElementNames.IsGreaterThan:
                    searchFilter = new IsGreaterThan();
                    break;
                case XmlElementNames.IsGreaterThanOrEqualTo:
                    searchFilter = new IsGreaterThanOrEqualTo();
                    break;
                case XmlElementNames.IsLessThan:
                    searchFilter = new IsLessThan();
                    break;
                case XmlElementNames.IsLessThanOrEqualTo:
                    searchFilter = new IsLessThanOrEqualTo();
                    break;
                default:
                    searchFilter = null;
                    break;
            }
            return searchFilter;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetXmlElementName();

        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonFilter = new JsonObject();
            jsonFilter.AddTypeParameter(this.GetXmlElementName());

            return jsonFilter;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteToXml(EwsServiceXmlWriter writer)
        {
            base.WriteToXml(writer, this.GetXmlElementName());
        }
    }
}