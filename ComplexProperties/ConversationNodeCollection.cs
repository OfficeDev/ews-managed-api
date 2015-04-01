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
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of conversation items.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ConversationNodeCollection : ComplexPropertyCollection<ConversationNode>, IJsonCollectionDeserializer
    {
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationNodeCollection"/> class.
        /// </summary>
        /// <param name="propertySet">The property set.</param>
        internal ConversationNodeCollection(PropertySet propertySet)
            : base()
        {
            this.propertySet = propertySet;
        }

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>ConversationItem.</returns>
        internal override ConversationNode CreateComplexProperty(string xmlElementName)
        {
            return new ConversationNode(this.propertySet);
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>ConversationItem.</returns>
        internal override ConversationNode CreateDefaultComplexProperty()
        {
            return new ConversationNode(this.propertySet);
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(ConversationNode complexProperty)
        {
            return complexProperty.GetXmlElementName();
        }

        #region IJsonCollectionDeserializer Members

        /// <summary>
        /// Loads from json collection.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.CreateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            foreach (object collectionEntry in jsonCollection)
            {
                JsonObject jsonEntry = collectionEntry as JsonObject;

                if (jsonEntry != null)
                {
                    ConversationNode node = new ConversationNode(this.propertySet);
                    node.LoadFromJson(jsonEntry, service);
                    this.InternalAdd(node);
                }
            }
        }

        /// <summary>
        /// Loads from json collection to update the existing collection element.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.UpdateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}