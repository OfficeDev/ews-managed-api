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
    using System.Linq;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents the response to a item search operation.
    /// </summary>
    /// <typeparam name="TItem">The type of items that the opeartion returned.</typeparam>
    internal sealed class FindItemResponse<TItem> : ServiceResponse
        where TItem : Item
    {
        private FindItemsResults<TItem> results;
        private bool isGrouped;
        private GroupedFindItemsResults<TItem> groupedFindResults;
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="FindItemResponse&lt;TItem&gt;"/> class.
        /// </summary>
        /// <param name="isGrouped">if set to <c>true</c> if grouped.</param>
        /// <param name="propertySet">The property set.</param>
        internal FindItemResponse(bool isGrouped, PropertySet propertySet)
            : base()
        {
            this.isGrouped = isGrouped;
            this.propertySet = propertySet;

            EwsUtilities.Assert(
                this.propertySet != null,
                "FindItemResponse.ctor",
                "PropertySet should not be null");
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RootFolder);

            int totalItemsInView = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
            bool moreItemsAvailable = !reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

            // Ignore IndexedPagingOffset attribute if moreItemsAvailable is false.
            int? nextPageOffset = moreItemsAvailable ? reader.ReadNullableAttributeValue<int>(XmlAttributeNames.IndexedPagingOffset) : null;

            if (!this.isGrouped)
            {
                this.results = new FindItemsResults<TItem>();
                this.results.TotalCount = totalItemsInView;
                this.results.NextPageOffset = nextPageOffset;
                this.results.MoreAvailable = moreItemsAvailable;
                InternalReadItemsFromXml(
                    reader,
                    this.propertySet,
                    this.results.Items);
            }
            else
            {
                this.groupedFindResults = new GroupedFindItemsResults<TItem>();
                this.groupedFindResults.TotalCount = totalItemsInView;
                this.groupedFindResults.NextPageOffset = nextPageOffset;
                this.groupedFindResults.MoreAvailable = moreItemsAvailable;

                reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Groups);

                if (!reader.IsEmptyElement) 
                {
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.GroupedItems))
                        {
                            string groupIndex = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.GroupIndex);

                            List<TItem> itemList = new List<TItem>();
                            InternalReadItemsFromXml(
                                reader,
                                this.propertySet,
                                itemList);

                            reader.ReadEndElement(XmlNamespace.Types, XmlElementNames.GroupedItems);

                            this.groupedFindResults.ItemGroups.Add(new ItemGroup<TItem>(groupIndex, itemList));
                        }
                    }
                    while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Groups));
                }
            }

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.RootFolder);

            reader.Read();

            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms) &&
                !reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        HighlightTerm term = new HighlightTerm();

                        term.LoadFromXml(
                            reader,
                            XmlNamespace.Types,
                            XmlElementNames.HighlightTerm);
                        this.results.HighlightTerms.Add(term);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.HighlightTerms));
            }
        }

        /// <summary>
        /// Read items from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertySet">The property set.</param>
        /// <param name="destinationList">The list in which to add the read items.</param>
        private static void InternalReadItemsFromXml(
            EwsServiceXmlReader reader,
            PropertySet propertySet,
            IList<TItem> destinationList)
        {
            EwsUtilities.Assert(
                destinationList != null,
                "FindItemResponse.InternalReadItemsFromXml",
                "destinationList is null.");

            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Items);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        TItem item = EwsUtilities.CreateEwsObjectFromXmlElementName<TItem>(reader.Service, reader.LocalName);

                        if (item == null)
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            item.LoadFromXml(
                                        reader,
                                        true, /* clearPropertyBag */
                                        propertySet,
                                        true  /* summaryPropertiesOnly */);

                            destinationList.Add(item);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Items));
            }
        }

        /// <summary>
        /// Creates an item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item</returns>
        private TItem CreateItemInstance(ExchangeService service, string xmlElementName)
        {
            return EwsUtilities.CreateEwsObjectFromXmlElementName<TItem>(service, xmlElementName);
        }

        /// <summary>
        /// Gets a grouped list of items matching the specified search criteria that were found in Exchange. ItemGroups is
        /// null if the search operation did not specify grouping options.
        /// </summary>
        public GroupedFindItemsResults<TItem> GroupedFindResults
        {
            get { return this.groupedFindResults; }
        }

        /// <summary>
        /// Gets the results of the search operation.
        /// </summary>
        public FindItemsResults<TItem> Results
        {
            get { return this.results; }
        }
    }
}