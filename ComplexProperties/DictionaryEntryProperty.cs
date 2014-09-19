// ---------------------------------------------------------------------------
// <copyright file="DictionaryEntryProperty.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DictionaryEntryProperty class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents an entry of a DictionaryProperty object.
    /// </summary>
    /// <remarks>
    /// All descendants of DictionaryEntryProperty must implement a parameterless
    /// constructor. That constructor does not have to be public.
    /// </remarks>
    /// <typeparam name="TKey">The type of the key used by this dictionary.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class DictionaryEntryProperty<TKey> : ComplexProperty
    {
        private TKey key;

        /// <summary>
        /// Initializes a new instance of the <see cref="DictionaryEntryProperty&lt;TKey&gt;"/> class.
        /// </summary>
        internal DictionaryEntryProperty()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DictionaryEntryProperty&lt;TKey&gt;"/> class.
        /// </summary>
        /// <param name="key">The key.</param>
        internal DictionaryEntryProperty(TKey key)
            : base()
        {
            this.key = key;
        }

        /// <summary>
        /// Gets or sets the key.
        /// </summary>
        /// <value>The key.</value>
        internal TKey Key
        {
            get { return this.key; }
            set { this.key = value; }
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.key = reader.ReadAttributeValue<TKey>(XmlAttributeNames.Key);
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Key, this.Key);
        }

        /// <summary>
        /// Writes the set update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="ownerDictionaryXmlElementName">Name of the owner dictionary XML element.</param>
        /// <returns>True if update XML was written.</returns>
        internal virtual bool WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            string ownerDictionaryXmlElementName)
        {
            return false;
        }

        /// <summary>
        /// Writes the set update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">The property definition.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        internal virtual bool WriteSetUpdateToJson(
             ExchangeService service,
             ServiceObject ewsObject,
             PropertyDefinition propertyDefinition,
             List<JsonObject> updates)
        {
            return false;
        }

        /// <summary>
        /// Writes the delete update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>True if update XML was written.</returns>
        internal virtual bool WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
        {
            return false;
        }

        /// <summary>
        /// Writes the delete update to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="updates">The updates.</param>
        /// <returns></returns>
        internal virtual bool WriteDeleteUpdateToJson(ExchangeService service, ServiceObject ewsObject, List<JsonObject> updates)
        {
            return false;
        }
    }
}
