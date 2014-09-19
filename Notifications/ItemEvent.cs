// ---------------------------------------------------------------------------
// <copyright file="ItemEvent.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemEvent class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents an event that applies to an item.
    /// </summary>
    public sealed class ItemEvent : NotificationEvent
    {
        /// <summary>
        /// Id of the item this event applies to.
        /// </summary>
        private ItemId itemId;

        /// <summary>
        /// Id of the item that moved or copied. This is only meaningful when EventType 
        /// is equal to either EventType.Moved or EventType.Copied. For all other event
        /// types, it's null.
        /// </summary>
        private ItemId oldItemId;

        /// <summary>
        /// Initializes a new instance of the <see cref="ItemEvent"/> class.
        /// </summary>
        /// <param name="eventType">Type of the event.</param>
        /// <param name="timestamp">The event timestamp.</param>
        internal ItemEvent(EventType eventType, DateTime timestamp)
            : base(eventType, timestamp)
        {
        }

        /// <summary>
        /// Load from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void InternalLoadFromXml(EwsServiceXmlReader reader)
        {
            base.InternalLoadFromXml(reader);

            this.itemId = new ItemId();
            this.itemId.LoadFromXml(reader, reader.LocalName);

            reader.Read();

            this.ParentFolderId = new FolderId();
            this.ParentFolderId.LoadFromXml(reader, XmlElementNames.ParentFolderId);

            switch (this.EventType)
            {
                case EventType.Moved:
                case EventType.Copied:
                    reader.Read();

                    this.oldItemId = new ItemId();
                    this.oldItemId.LoadFromXml(reader, reader.LocalName);

                    reader.Read();

                    this.OldParentFolderId = new FolderId();
                    this.OldParentFolderId.LoadFromXml(reader, reader.LocalName);
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonEvent">The json event.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonEvent, ExchangeService service)
        {
            this.itemId = new ItemId();
            this.itemId.LoadFromJson(jsonEvent.ReadAsJsonObject(XmlElementNames.ItemId), service);

            this.ParentFolderId = new FolderId();
            this.ParentFolderId.LoadFromJson(jsonEvent.ReadAsJsonObject(XmlElementNames.ParentFolderId), service);

            switch (this.EventType)
            {
                case EventType.Moved:
                case EventType.Copied:

                    this.oldItemId = new ItemId();
                    this.oldItemId.LoadFromJson(jsonEvent.ReadAsJsonObject(JsonNames.OldItemId), service);

                    this.OldParentFolderId = new FolderId();
                    this.OldParentFolderId.LoadFromJson(jsonEvent.ReadAsJsonObject(XmlElementNames.OldParentFolderId), service);
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Gets the Id of the item this event applies to.
        /// </summary>
        public ItemId ItemId
        {
            get { return this.itemId; }
        }

        /// <summary>
        /// Gets the Id of the item that was moved or copied. OldItemId is only meaningful
        /// when EventType is equal to either EventType.Moved or EventType.Copied. For
        /// all other event types, OldItemId is null.
        /// </summary>
        public ItemId OldItemId
        {
            get { return this.oldItemId; }
        }
    }
}
