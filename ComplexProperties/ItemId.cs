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
    /// Represents the Id of an Exchange item.
    /// </summary>
    public class ItemId : ServiceId
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ItemId"/> class.
        /// </summary>
        internal ItemId()
            : base()
        {
        }

        /// <summary>
        /// Defines an implicit conversion between string and ItemId.
        /// </summary>
        /// <param name="uniqueId">The unique Id to convert to ItemId.</param>
        /// <returns>An ItemId initialized with the specified unique Id.</returns>
        public static implicit operator ItemId(string uniqueId)
        {
            return new ItemId(uniqueId);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.ItemId;
        }

        /// <summary>
        /// Initializes a new instance of ItemId.
        /// </summary>
        /// <param name="uniqueId">The unique Id used to initialize the ItemId.</param>
        public ItemId(string uniqueId)
            : base(uniqueId)
        {
        }
    }
}