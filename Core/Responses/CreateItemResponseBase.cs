// ---------------------------------------------------------------------------
// <copyright file="CreateItemResponseBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateItemResponseBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the base response class for item creation operations.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal abstract class CreateItemResponseBase : ServiceResponse
    {
        private List<Item> items;

        /// <summary>
        /// Gets Item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item.</returns>
        internal abstract Item GetObjectInstance(ExchangeService service, string xmlElementName);

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateItemResponseBase"/> class.
        /// </summary>
        internal CreateItemResponseBase()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.items = reader.ReadServiceObjectsCollectionFromXml<Item>(
                XmlElementNames.Items,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            this.items = new EwsServiceJsonReader(service).ReadServiceObjectsCollectionFromJson<Item>(
                responseObject,
                XmlElementNames.Items,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */
        }

        /// <summary>
        /// Gets the items.
        /// </summary>
        public List<Item> Items
        {
            get { return this.items; }
        }
    }
}
