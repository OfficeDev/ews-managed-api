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
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents the base view class for search operations.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class ViewBase
    {
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewBase"/> class.
        /// </summary>
        internal ViewBase()
        {
        }

        /// <summary>
        /// Validates this view.
        /// </summary>
        /// <param name="request">The request using this view.</param>
        internal virtual void InternalValidate(ServiceRequestBase request)
        {
            if (this.PropertySet != null)
            {
                this.PropertySet.InternalValidate();
                this.PropertySet.ValidateForRequest(request, true /*summaryPropertiesOnly*/);
            }
        }

        /// <summary>
        /// Writes this view to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void InternalWriteViewToXml(EwsServiceXmlWriter writer)
        {
            int? maxEntriesReturned = this.GetMaxEntriesReturned();

            if (maxEntriesReturned.HasValue)
            {
                writer.WriteAttributeValue(XmlAttributeNames.MaxEntriesReturned, maxEntriesReturned.Value);
            }
        }

        /// <summary>
        /// Writes the search settings to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="groupBy">The group by clause.</param>
        internal abstract void InternalWriteSearchSettingsToXml(EwsServiceXmlWriter writer, Grouping groupBy);

        /// <summary>
        /// Writes OrderBy property to XML.
        /// </summary>
        /// <param name="writer">The writer</param>
        internal abstract void WriteOrderByToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Gets the name of the view XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal abstract string GetViewXmlElementName();

        /// <summary>
        /// Gets the maximum number of items or folders the search operation should return.
        /// </summary>
        /// <returns>The maximum number of items or folders that should be returned by the search operation.</returns>
        internal abstract int? GetMaxEntriesReturned();

        /// <summary>
        /// Gets the type of service object this view applies to.
        /// </summary>
        /// <returns>A ServiceObjectType value.</returns>
        internal abstract ServiceObjectType GetServiceObjectType();

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteAttributesToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="groupBy">The group by clause.</param>
        internal virtual void WriteToXml(EwsServiceXmlWriter writer, Grouping groupBy)
        {
            this.GetPropertySetOrDefault().WriteToXml(writer, this.GetServiceObjectType());

            writer.WriteStartElement(XmlNamespace.Messages, this.GetViewXmlElementName());

            this.InternalWriteViewToXml(writer);

            writer.WriteEndElement(); // this.GetViewXmlElementName()

            this.InternalWriteSearchSettingsToXml(writer, groupBy);
        }

        /// <summary>
        /// Gets the property set or the default.
        /// </summary>
        /// <returns>PropertySet</returns>
        internal PropertySet GetPropertySetOrDefault()
        {
            // If property set is null, default is FirstClassProperties
            return this.PropertySet ?? PropertySet.FirstClassProperties;
        }

        /// <summary>
        /// Gets or sets the property set. PropertySet determines which properties will be loaded on found items. If PropertySet is null,
        /// all first class properties are loaded on found items.
        /// </summary>
        public PropertySet PropertySet
        {
            get { return this.propertySet; }
            set { this.propertySet = value; }
        }
    }
}