// ---------------------------------------------------------------------------
// <copyright file="SearchFilter.Not.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the Not class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <content>
    /// Contains nested type SearchFilter.Not.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a search filter that negates another. Applications can use NotFilter to define
        /// conditions such as "NOT(other filter)".
        /// </summary>
        public sealed class Not : SearchFilter
        {
            private SearchFilter searchFilter;

            /// <summary>
            /// Initializes a new instance of the <see cref="Not"/> class.
            /// </summary>
            public Not()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="Not"/> class.
            /// </summary>
            /// <param name="searchFilter">The search filter to negate. Available search filter classes include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.</param>
            public Not(SearchFilter searchFilter)
                : base()
            {
                this.searchFilter = searchFilter;
            }

            /// <summary>
            /// A search filter has changed.
            /// </summary>
            /// <param name="complexProperty">The complex property.</param>
            private void SearchFilterChanged(ComplexProperty complexProperty)
            {
                this.Changed();
            }

            /// <summary>
            /// Validate instance.
            /// </summary>
            internal override void InternalValidate()
            {
                if (this.searchFilter == null)
                {
                    throw new ServiceValidationException(Strings.SearchFilterMustBeSet);
                }
            }

            /// <summary>
            /// Gets the name of the XML element.
            /// </summary>
            /// <returns>XML element name.</returns>
            internal override string GetXmlElementName()
            {
                return XmlElementNames.Not;
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if element was read.</returns>
            internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
                this.searchFilter = SearchFilter.LoadFromXml(reader);
                return true;
            }

            internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
            {
                this.searchFilter = SearchFilter.LoadSearchFilterFromJson(jsonProperty.ReadAsJsonObject(XmlElementNames.Item), service);
            }

            /// <summary>
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                this.SearchFilter.WriteToXml(writer);
            }

            internal override object InternalToJson(ExchangeService service)
            {
                JsonObject jsonFilter = base.InternalToJson(service) as JsonObject;

                jsonFilter.Add(XmlElementNames.Item, this.SearchFilter.InternalToJson(service));

                return jsonFilter;
            }

            /// <summary>
            /// Gets or sets the search filter to negate. Available search filter classes include
            /// SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.
            /// </summary>
            public SearchFilter SearchFilter
            {
                get
                {
                    return this.searchFilter;
                }

                set
                {
                    if (this.searchFilter != null)
                    {
                        this.searchFilter.OnChange -= this.SearchFilterChanged;
                    }

                    this.SetFieldValue<SearchFilter>(ref this.searchFilter, value);

                    if (this.searchFilter != null)
                    {
                        this.searchFilter.OnChange += this.SearchFilterChanged;
                    }
                }
            }
        }
    }
}