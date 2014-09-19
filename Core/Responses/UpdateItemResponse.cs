// ---------------------------------------------------------------------------
// <copyright file="UpdateItemResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UpdateItemResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the response to an individual item update operation.
    /// </summary>
    public sealed class UpdateItemResponse : ServiceResponse
    {
        private Item item;
        private Item returnedItem;
        private int conflictCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateItemResponse"/> class.
        /// </summary>
        /// <param name="item">The item.</param>
        internal UpdateItemResponse(Item item)
            : base()
        {
            EwsUtilities.Assert(
                item != null,
                "UpdateItemResponse.ctor",
                "item is null");

            this.item = item;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadServiceObjectsCollectionFromXml<Item>(
                XmlElementNames.Items,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */

            // ConflictResults was only added in 2007 SP1 so if this was a 2007 RTM request we shouldn't expect to find the element
            if (!reader.Service.Exchange2007CompatibilityMode)
            {
                reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ConflictResults);
                this.conflictCount = reader.ReadElementValue<int>(XmlNamespace.Types, XmlElementNames.Count);
                reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.ConflictResults);
            }

            // If UpdateItem returned an item that has the same Id as the item that
            // is being updated, this is a "normal" UpdateItem operation, and we need
            // to update the ChangeKey of the item being updated with the one that was
            // returned. Also set returnedItem to indicate that no new item was returned.
            //
            // Otherwise, this in a "special" UpdateItem operation, such as a recurring
            // task marked as complete (the returned item in that case is the one-off
            // task that represents the completed instance).
            //
            // Note that there can be no returned item at all, as in an UpdateItem call
            // with MessageDisposition set to SendOnly or SendAndSaveCopy.
            if (this.returnedItem != null)
            {
                if (this.item.Id.UniqueId == this.returnedItem.Id.UniqueId)
                {
                    this.item.Id.ChangeKey = this.returnedItem.Id.ChangeKey;
                    this.returnedItem = null;
                }
            }
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            new EwsServiceJsonReader(service).ReadServiceObjectsCollectionFromJson<Item>(
                responseObject,
                XmlElementNames.Items,
                this.GetObjectInstance,
                false,  /* clearPropertyBag */
                null,   /* requestedPropertySet */
                false); /* summaryPropertiesOnly */

            // ConflictResults was only added in 2007 SP1 so if this was a 2007 RTM request we shouldn't expect to find the element
            if (!service.Exchange2007CompatibilityMode)
            {
                this.conflictCount = responseObject.ReadAsJsonObject(XmlElementNames.ConflictResults).ReadAsInt(XmlElementNames.Count);
            }

            // If UpdateItem returned an item that has the same Id as the item that
            // is being updated, this is a "normal" UpdateItem operation, and we need
            // to update the ChangeKey of the item being updated with the one that was
            // returned. Also set returnedItem to indicate that no new item was returned.
            //
            // Otherwise, this in a "special" UpdateItem operation, such as a recurring
            // task marked as complete (the returned item in that case is the one-off
            // task that represents the completed instance).
            //
            // Note that there can be no returned item at all, as in an UpdateItem call
            // with MessageDisposition set to SendOnly or SendAndSaveCopy.
            if (this.returnedItem != null)
            {
                if (this.item.Id.UniqueId == this.returnedItem.Id.UniqueId)
                {
                    this.item.Id.ChangeKey = this.returnedItem.Id.ChangeKey;
                    this.returnedItem = null;
                }
            }
        }

        /// <summary>
        /// Clears the change log of the created folder if the creation succeeded.
        /// </summary>
        internal override void Loaded()
        {
            if (this.Result == ServiceResult.Success)
            {
                this.item.ClearChangeLog();
            }
        }

        /// <summary>
        /// Gets Item instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Item.</returns>
        private Item GetObjectInstance(ExchangeService service, string xmlElementName)
        {
            this.returnedItem = EwsUtilities.CreateEwsObjectFromXmlElementName<Item>(service, xmlElementName);

            return this.returnedItem;
        }

        /// <summary>
        /// Gets the item that was returned by the update operation. ReturnedItem is set only when a recurring Task
        /// is marked as complete or when its recurrence pattern changes. 
        /// </summary>
        public Item ReturnedItem
        {
            get { return this.returnedItem; }
        }

        /// <summary>
        /// Gets the number of property conflicts that were resolved during the update operation.
        /// </summary>
        public int ConflictCount
        {
            get { return this.conflictCount; }
        }
    }
}
