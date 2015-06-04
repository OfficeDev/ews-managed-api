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

    /// <content>
    /// Contains nested type SearchFilter.SearchFilterCollection.
    /// </content>
    public abstract partial class SearchFilter
    {
        /// <summary>
        /// Represents a collection of search filters linked by a logical operator. Applications can
        /// use SearchFilterCollection to define complex search filters such as "Condition1 AND Condition2".
        /// </summary>
        public sealed class SearchFilterCollection : SearchFilter, IEnumerable<SearchFilter>
        {
            private List<SearchFilter> searchFilters = new List<SearchFilter>();
            private LogicalOperator logicalOperator = LogicalOperator.And;

            /// <summary>
            /// Initializes a new instance of the <see cref="SearchFilterCollection"/> class.
            /// The LogicalOperator property is initialized to LogicalOperator.And.
            /// </summary>
            public SearchFilterCollection()
                : base()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="SearchFilterCollection"/> class.
            /// </summary>
            /// <param name="logicalOperator">The logical operator used to initialize the collection.</param>
            public SearchFilterCollection(LogicalOperator logicalOperator)
                : base()
            {
                this.logicalOperator = logicalOperator;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="SearchFilterCollection"/> class.
            /// </summary>
            /// <param name="logicalOperator">The logical operator used to initialize the collection.</param>
            /// <param name="searchFilters">The search filters to add to the collection.</param>
            public SearchFilterCollection(LogicalOperator logicalOperator, params SearchFilter[] searchFilters)
                : this(logicalOperator)
            {
                this.AddRange(searchFilters);
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="SearchFilterCollection"/> class.
            /// </summary>
            /// <param name="logicalOperator">The logical operator used to initialize the collection.</param>
            /// <param name="searchFilters">The search filters to add to the collection.</param>
            public SearchFilterCollection(LogicalOperator logicalOperator, IEnumerable<SearchFilter> searchFilters)
                : this(logicalOperator)
            {
                this.AddRange(searchFilters);
            }

            /// <summary>
            /// Validate instance.
            /// </summary>
            internal override void InternalValidate()
            {
                for (int i = 0; i < this.Count; i++)
                {
                    try
                    {
                        this[i].InternalValidate();
                    }
                    catch (ServiceValidationException e)
                    {
                        throw new ServiceValidationException(string.Format(Strings.SearchFilterAtIndexIsInvalid, i), e);
                    }
                }
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
            /// Gets the name of the XML element.
            /// </summary>
            /// <returns>XML element name.</returns>
            internal override string GetXmlElementName()
            {
                return this.LogicalOperator.ToString();
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if element was read.</returns>
            internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
                this.Add(SearchFilter.LoadFromXml(reader));
                return true;
            }

            /// <summary>
            /// Writes the elements to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
                foreach (SearchFilter searchFilter in this)
                {
                    searchFilter.WriteToXml(writer);
                }
            }

            /// <summary>
            /// Writes to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void WriteToXml(EwsServiceXmlWriter writer)
            {
                // If there is only one filter in the collection, which developers tend to do,
                // we need to not emit the collection and instead only emit the one filter within
                // the collection. This is to work around the fact that EWS does not allow filter
                // collections that have less than two elements.
                if (this.Count == 1)
                {
                    this[0].WriteToXml(writer);
                }
                else
                {
                    base.WriteToXml(writer);
                }
            }

            /// <summary>
            /// Adds a search filter of any type to the collection.
            /// </summary>
            /// <param name="searchFilter">The search filter to add. Available search filter classes include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.</param>
            public void Add(SearchFilter searchFilter)
            {
                if (searchFilter == null)
                {
                    throw new ArgumentNullException("searchFilter");
                }

                searchFilter.OnChange += this.SearchFilterChanged;
                this.searchFilters.Add(searchFilter);
                this.Changed();
            }

            /// <summary>
            /// Adds multiple search filters to the collection.
            /// </summary>
            /// <param name="searchFilters">The search filters to add. Available search filter classes include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.</param>
            public void AddRange(IEnumerable<SearchFilter> searchFilters)
            {
                if (searchFilters == null)
                {
                    throw new ArgumentNullException("searchFilters");
                }

                foreach (SearchFilter searchFilter in searchFilters)
                {
                    searchFilter.OnChange += this.SearchFilterChanged;
                }
                this.searchFilters.AddRange(searchFilters);
                this.Changed();
            }

            /// <summary>
            /// Clears the collection.
            /// </summary>
            public void Clear()
            {
                if (this.Count > 0)
                {
                    foreach (SearchFilter searchFilter in this)
                    {
                        searchFilter.OnChange -= this.SearchFilterChanged;
                    }
                    this.searchFilters.Clear();
                    this.Changed();
                }
            }

            /// <summary>
            /// Determines whether a specific search filter is in the collection.
            /// </summary>
            /// <param name="searchFilter">The search filter to locate in the collection.</param>
            /// <returns>True is the search filter was found in the collection, false otherwise.</returns>
            public bool Contains(SearchFilter searchFilter)
            {
                return this.searchFilters.Contains(searchFilter);
            }

            /// <summary>
            /// Removes a search filter from the collection.
            /// </summary>
            /// <param name="searchFilter">The search filter to remove.</param>
            public void Remove(SearchFilter searchFilter)
            {
                if (searchFilter == null)
                {
                    throw new ArgumentNullException("searchFilter");
                }

                if (this.Contains(searchFilter))
                {
                    searchFilter.OnChange -= this.SearchFilterChanged;
                    this.searchFilters.Remove(searchFilter);
                    this.Changed();
                }
            }

            /// <summary>
            /// Removes the search filter at the specified index from the collection.
            /// </summary>
            /// <param name="index">The zero-based index of the search filter to remove.</param>
            public void RemoveAt(int index)
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                this[index].OnChange -= this.SearchFilterChanged;
                this.searchFilters.RemoveAt(index);
                this.Changed();
            }

            /// <summary>
            /// Gets the total number of search filters in the collection.
            /// </summary>
            public int Count
            {
                get { return this.searchFilters.Count; }
            }

            /// <summary>
            /// Gets or sets the search filter at the specified index.
            /// </summary>
            /// <param name="index">The zero-based index of the search filter to get or set.</param>
            /// <returns>The search filter at the specified index.</returns>
            public SearchFilter this[int index]
            {
                get
                {
                    if (index < 0 || index >= this.Count)
                    {
                        throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                    }

                    return this.searchFilters[index];
                }

                set
                {
                    if (index < 0 || index >= this.Count)
                    {
                        throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                    }

                    this.searchFilters[index] = value;
                }
            }

            /// <summary>
            /// Gets or sets the logical operator that links the serach filters in this collection.
            /// </summary>
            public LogicalOperator LogicalOperator
            {
                get { return this.logicalOperator; }
                set { this.logicalOperator = value; }
            }

            #region IEnumerable<SearchCondition> Members

            /// <summary>
            /// Gets an enumerator that iterates through the elements of the collection.
            /// </summary>
            /// <returns>An IEnumerator for the collection.</returns>
            public IEnumerator<SearchFilter> GetEnumerator()
            {
                return this.searchFilters.GetEnumerator();
            }

            #endregion

            #region IEnumerable Members

            /// <summary>
            /// Gets an enumerator that iterates through the elements of the collection.
            /// </summary>
            /// <returns>An IEnumerator for the collection.</returns>
            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return this.searchFilters.GetEnumerator();
            }

            #endregion
        }
    }
}