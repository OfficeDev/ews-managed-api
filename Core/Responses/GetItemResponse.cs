// ---------------------------------------------------------------------------
// <copyright file="GetItemResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetItemResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a response to an individual item retrieval operation.
    /// </summary>
    public sealed class GetItemResponse : ServiceResponse
    {
        private Item item;
        private PropertySet propertySet;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetItemResponse"/> class.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="propertySet">The property set.</param>
        internal GetItemResponse(Item item, PropertySet propertySet)
            : base()
        {
            this.item = item;
            this.propertySet = propertySet;

            EwsUtilities.Assert(
                this.propertySet != null,
                "GetItemResponse.ctor",
                "PropertySet should not be null");
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            List<Item> items = reader.ReadServiceObjectsCollectionFromXml<Item>(
                XmlElementNames.Items,
                this.GetObjectInstance,
                true,               /* clearPropertyBag */
                this.propertySet,   /* requestedPropertySet */
                false);             /* summaryPropertiesOnly */

            this.item = items[0];
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            List<Item> items = new EwsServiceJsonReader(service).ReadServiceObjectsCollectionFromJson<Item>(
                responseObject,
                XmlElementNames.Items,
                this.GetObjectInstance,
                true,               /* clearPropertyBag */
                this.propertySet,   /* requestedPropertySet */
                false);             /* summaryPropertiesOnly */

            this.item = items[0];
        }

        /// <summary>
        /// Gets Item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item.</returns>
        private Item GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            if (this.Item != null)
            {
                return this.Item;
            }
            else
            {
                return EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);
            }
        }

        /// <summary>
        /// Gets the item that was retrieved.
        /// </summary>
        public Item Item
        {
            get { return this.item; }
        }
    }
}
