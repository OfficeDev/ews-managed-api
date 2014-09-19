// ---------------------------------------------------------------------------
// <copyright file="ConversationNode.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationNode class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the response to a GetConversationItems operation.
    /// </summary>
    public sealed class ConversationNode : ComplexProperty
    {
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationNode"/> class.
        /// </summary>
        /// <param name="propertySet">The property set.</param>
        internal ConversationNode(PropertySet propertySet)
            : base()
        {
            this.propertySet = propertySet;
        }

        /// <summary>
        /// Gets or sets the Internet message id of the node.
        /// </summary>
        public string InternetMessageId { get; set; }

        /// <summary>
        /// Gets or sets the Internet message id of the parent node.
        /// </summary>
        public string ParentInternetMessageId { get; set; }

        /// <summary>
        /// Gets or sets the items.
        /// </summary>
        public List<Item> Items { get; set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.InternetMessageId:
                    this.InternetMessageId = reader.ReadElementValue();
                    return true;

                case XmlElementNames.ParentInternetMessageId:
                    this.ParentInternetMessageId = reader.ReadElementValue();
                    return true;

                case XmlElementNames.Items:
                    this.Items = reader.ReadServiceObjectsCollectionFromXml<Item>(
                                        XmlNamespace.Types,
                                        XmlElementNames.Items,
                                        this.GetObjectInstance,
                                        true,               /* clearPropertyBag */
                                        this.propertySet,   /* requestedPropertySet */
                                        false);             /* summaryPropertiesOnly */
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            this.InternetMessageId = jsonProperty.ReadAsString(XmlElementNames.ConversationIndex);

            if (jsonProperty.ContainsKey(XmlElementNames.ParentInternetMessageId))
            {
                this.ParentInternetMessageId = jsonProperty.ReadAsString(XmlElementNames.ParentInternetMessageId);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.Items))
            {
                EwsServiceJsonReader jsonReader = new EwsServiceJsonReader(service);

                this.Items = jsonReader.ReadServiceObjectsCollectionFromJson<Item>(
                    jsonProperty,
                    XmlElementNames.Items,
                    this.GetObjectInstance,
                    false,              /* clearPropertyBag */
                    this.propertySet,   /* requestedPropertySet */
                    false);             /* summaryPropertiesOnly */
            }
        }

        /// <summary>
        /// Gets the item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item.</returns>
        private Item GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            return EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal string GetXmlElementName()
        {
            return XmlElementNames.ConversationNode;
        }
    }
}
