// ---------------------------------------------------------------------------
// <copyright file="FolderEvent.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FolderEvent class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents an event that applies to a folder.
    /// </summary>
    public class FolderEvent : NotificationEvent
    {
        private FolderId folderId;
        private FolderId oldFolderId;

        /// <summary>
        /// The new number of unread messages. This is is only meaningful when EventType
        /// is equal to EventType.Modified. For all other event types, it's null.
        /// </summary>
        private int? unreadCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderEvent"/> class.
        /// </summary>
        /// <param name="eventType">Type of the event.</param>
        /// <param name="timestamp">The event timestamp.</param>
        internal FolderEvent(EventType eventType, DateTime timestamp)
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

            this.folderId = new FolderId();
            this.folderId.LoadFromXml(reader, reader.LocalName);

            reader.Read();

            this.ParentFolderId = new FolderId();
            this.ParentFolderId.LoadFromXml(reader, XmlElementNames.ParentFolderId);

            switch (this.EventType)
            {
                case EventType.Moved:
                case EventType.Copied:
                    reader.Read();

                    this.oldFolderId = new FolderId();
                    this.oldFolderId.LoadFromXml(reader, reader.LocalName);

                    reader.Read();

                    this.OldParentFolderId = new FolderId();
                    this.OldParentFolderId.LoadFromXml(reader, reader.LocalName);
                    break;

                case EventType.Modified:
                    reader.Read();
                    if (reader.IsStartElement())
                    {
                        reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, XmlElementNames.UnreadCount);
                        this.unreadCount = int.Parse(reader.ReadValue());
                    }
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
            this.folderId = new FolderId();
            this.folderId.LoadFromJson(jsonEvent.ReadAsJsonObject(XmlElementNames.FolderId), service);

            this.ParentFolderId = new FolderId();
            this.ParentFolderId.LoadFromJson(jsonEvent.ReadAsJsonObject(XmlElementNames.ParentFolderId), service);

            switch (this.EventType)
            {
                case EventType.Moved:
                case EventType.Copied:

                    this.oldFolderId = new FolderId();
                    this.oldFolderId.LoadFromJson(jsonEvent.ReadAsJsonObject(JsonNames.OldFolderId), service);

                    this.OldParentFolderId = new FolderId();
                    this.OldParentFolderId.LoadFromJson(jsonEvent.ReadAsJsonObject(XmlElementNames.OldParentFolderId), service);
                    break;

                case EventType.Modified:
                    if (jsonEvent.ContainsKey(XmlElementNames.UnreadCount))
                    {
                        this.unreadCount = jsonEvent.ReadAsInt(XmlElementNames.UnreadCount);
                    }
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Gets the Id of the folder this event applies to.
        /// </summary>
        public FolderId FolderId
        {
            get { return this.folderId; }
        }

        /// <summary>
        /// Gets the Id of the folder that was moved or copied. OldFolderId is only meaningful
        /// when EventType is equal to either EventType.Moved or EventType.Copied. For all
        /// other event types, OldFolderId is null.
        /// </summary>
        public FolderId OldFolderId
        {
            get { return this.oldFolderId; }
        }

        /// <summary>
        /// Gets the new number of unread messages. This is is only meaningful when 
        /// EventType is equal to EventType.Modified. For all other event types, 
        /// UnreadCount is null.
        /// </summary>
        public int? UnreadCount
        {
            get { return this.unreadCount; }
        }
    }
}
