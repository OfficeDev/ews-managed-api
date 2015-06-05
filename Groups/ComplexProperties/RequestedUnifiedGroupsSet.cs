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

namespace Microsoft.Exchange.WebServices.Data.Groups
{
    /// <summary>
    /// Defines the RequestedUnifiedGroupsSet class.
    /// </summary>
    public sealed class RequestedUnifiedGroupsSet : ComplexProperty, ISelfValidate
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedUnifiedGroupsSet"/> class.
        /// </summary>
        public RequestedUnifiedGroupsSet()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedUnifiedGroupsSet"/> class.
        /// </summary>
        /// <param name="filterType">The filterType for the list of groups to be returned</param>
        /// <param name="sortType">The sort type for the list of groups to be returned</param>
        /// <param name="sortDirection">The sort direction for list of groups to be returned</param>
        public RequestedUnifiedGroupsSet(
            UnifiedGroupsFilterType filterType,
            UnifiedGroupsSortType sortType,
            SortDirection sortDirection)
        {
            this.FilterType = filterType;
            this.SortType = sortType;
            this.SortDirection = sortDirection;
        }

        /// <summary>
        /// Gets or sets the sort type for the list of groups to be returned
        /// </summary>
        public UnifiedGroupsSortType SortType { get; set; }

        /// <summary>
        /// Gets or sets the filter Type for the list of groups to be returned
        /// </summary>
        public UnifiedGroupsFilterType FilterType { get; set; }

        /// <summary>
        /// Gets or sets the Sort Direction for the list of groups to be returned.
        /// </summary>
        public SortDirection SortDirection { get; set; }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SortType, this.SortType.ToString());
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FilterType, this.FilterType.ToString());
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SortDirection, this.SortDirection.ToString());

            writer.WriteEndElement();
        }
    }
}