// ---------------------------------------------------------------------------
// <copyright file="ConversationNodeCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationNodeCollection class.</summary>
//-----------------------------------------------------------------------

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
