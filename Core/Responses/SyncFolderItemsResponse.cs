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
// <summary>Defines the SyncFolderItemsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to a folder items synchronization operation.
    /// </summary>
    public sealed class SyncFolderItemsResponse : SyncResponse<Item, ItemChange>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncFolderItemsResponse"/> class.
        /// </summary>
        /// <param name="propertySet">PropertySet from request.</param>
        internal SyncFolderItemsResponse(PropertySet propertySet)
            : base(propertySet)
        {
        }

        /// <summary>
        /// Gets the name of the includes last in range XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetIncludesLastInRangeXmlElementName()
        {
            return XmlElementNames.IncludesLastItemInRange;
        }

        /// <summary>
        /// Creates an item change instance.
        /// </summary>
        /// <returns>ItemChange instance</returns>
        internal override ItemChange CreateChangeInstance()
        {
            return new ItemChange();
        }

        /// <summary>
        /// Gets the name of the change element.
        /// </summary>
        /// <returns>Change element name.</returns>
        internal override string GetChangeElementName()
        {
            return XmlElementNames.Item;
        }

        /// <summary>
        /// Gets the name of the change id element.
        /// </summary>
        /// <returns>Change id element name.</returns>
        internal override string GetChangeIdElementName()
        {
            return XmlElementNames.ItemId;
        }

        /// <summary>
        /// Gets a value indicating whether this request returns full or summary properties.
        /// </summary>
        /// <value>
        /// <c>true</c> if summary properties only; otherwise, <c>false</c>.
        /// </value>
        internal override bool SummaryPropertiesOnly
        {
            get { return true; }
        }
    }
}
