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
    using System.Collections.ObjectModel;
    using System.Linq;

    /// <summary>
    /// Represents a collection of notification events.
    /// </summary>
    public sealed class GetEventsResults
    {
        /// <summary>
        /// Map XML element name to notification event type.
        /// </summary>
        /// <remarks>
        /// If you add a new notification event type, you'll need to add a new entry to the dictionary here.
        /// </remarks>
        private static LazyMember<Dictionary<string, EventType>> xmlElementNameToEventTypeMap = new LazyMember<Dictionary<string, EventType>>(
            delegate()
            {
                Dictionary<string, EventType> result = new Dictionary<string, EventType>();

                result.Add(XmlElementNames.CopiedEvent, EventType.Copied);
                result.Add(XmlElementNames.CreatedEvent, EventType.Created);
                result.Add(XmlElementNames.DeletedEvent, EventType.Deleted);
                result.Add(XmlElementNames.ModifiedEvent, EventType.Modified);
                result.Add(XmlElementNames.MovedEvent, EventType.Moved);
                result.Add(XmlElementNames.NewMailEvent, EventType.NewMail);
                result.Add(XmlElementNames.StatusEvent, EventType.Status);
                result.Add(XmlElementNames.FreeBusyChangedEvent, EventType.FreeBusyChanged);

                return result;
            });

        /// <summary>
        /// Gets the XML element name to event type mapping.
        /// </summary>
        /// <value>The XML element name to event type mapping.</value>
        internal static Dictionary<string, EventType> XmlElementNameToEventTypeMap
        {
            get
            {
                return GetEventsResults.xmlElementNameToEventTypeMap.Member;
            }
        }

        /// <summary>
        /// Watermark in event.
        /// </summary>
        private string newWatermark;

        /// <summary>
        /// Subscription id.
        /// </summary>
        private string subscriptionId;

        /// <summary>
        /// Previous watermark.
        /// </summary>
        private string previousWatermark;

        /// <summary>
        /// True if more events available for this subscription.
        /// </summary>
        private bool moreEventsAvailable;

        /// <summary>
        /// Collection of notification events.
        /// </summary>
        private Collection<NotificationEvent> events = new Collection<NotificationEvent>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetEventsResults"/> class.
        /// </summary>
        internal GetEventsResults()
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Notification);

            this.subscriptionId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SubscriptionId);
            this.previousWatermark = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.PreviousWatermark);
            this.moreEventsAvailable = reader.ReadElementValue<bool>(XmlNamespace.Types, XmlElementNames.MoreEvents);

            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    string eventElementName = reader.LocalName;
                    EventType eventType;

                    if (xmlElementNameToEventTypeMap.Member.TryGetValue(eventElementName, out eventType))
                    {
                        this.newWatermark = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.Watermark);

                        if (eventType == EventType.Status)
                        {
                            // We don't need to return status events
                            reader.ReadEndElementIfNecessary(XmlNamespace.Types, eventElementName);
                        }
                        else
                        {
                            this.LoadNotificationEventFromXml(
                                reader,
                                eventElementName,
                                eventType);
                        }
                    }
                    else
                    {
                        reader.SkipCurrentElement();
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Notification));
        }

        /// <summary>
        /// Loads a notification event from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="eventElementName">Name of the event XML element.</param>
        /// <param name="eventType">Type of the event.</param>
        private void LoadNotificationEventFromXml(
            EwsServiceXmlReader reader,
            string eventElementName,
            EventType eventType)
        {
            DateTime timestamp = reader.ReadElementValue<DateTime>(XmlNamespace.Types, XmlElementNames.TimeStamp);

            NotificationEvent notificationEvent;

            reader.Read();

            if (reader.LocalName == XmlElementNames.FolderId)
            {
                notificationEvent = new FolderEvent(eventType, timestamp);
            }
            else
            {
                notificationEvent = new ItemEvent(eventType, timestamp);
            }

            notificationEvent.LoadFromXml(reader, eventElementName);
            this.events.Add(notificationEvent);
        }

        /// <summary>
        /// Gets the Id of the subscription the collection is associated with.
        /// </summary>
        internal string SubscriptionId
        {
            get { return this.subscriptionId; }
        }

        /// <summary>
        /// Gets the subscription's previous watermark.
        /// </summary>
        internal string PreviousWatermark
        {
            get { return this.previousWatermark; }
        }

        /// <summary>
        /// Gets the subscription's new watermark.
        /// </summary>
        internal string NewWatermark
        {
            get { return this.newWatermark; }
        }

        /// <summary>
        /// Gets a value indicating whether more events are available on the Exchange server.
        /// </summary>
        internal bool MoreEventsAvailable
        {
            get { return this.moreEventsAvailable; }
        }

        /// <summary>
        /// Gets the collection of folder events.
        /// </summary>
        /// <value>The folder events.</value>
        public IEnumerable<FolderEvent> FolderEvents
        {
            get { return this.events.OfType<FolderEvent>(); }
        }

        /// <summary>
        /// Gets the collection of item events.
        /// </summary>
        /// <value>The item events.</value>
        public IEnumerable<ItemEvent> ItemEvents
        {
            get { return this.events.OfType<ItemEvent>(); }
        }

        /// <summary>
        /// Gets the collection of all events.
        /// </summary>
        /// <value>The events.</value>
        public Collection<NotificationEvent> AllEvents
        {
            get { return this.events; }
        }
    }
}