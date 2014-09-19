// ---------------------------------------------------------------------------
// <copyright file="GetStreamingEventsResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetStreamingEventsResults class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;

    /// <summary>
    /// Represents a collection of notification events.
    /// </summary>
    internal sealed class GetStreamingEventsResults
    {        
        /// <summary>
        /// Structure to track a subscription and its associated notification events.
        /// </summary>
        internal struct NotificationGroup
        {
            /// <summary>
            /// Subscription Id
            /// </summary>
            internal string SubscriptionId;

            /// <summary>
            /// Events in the response associated with the subscription id.
            /// </summary>
            internal Collection<NotificationEvent> Events;
        }

        /// <summary>
        /// Collection of notification events.
        /// </summary>
        private Collection<NotificationGroup> events = new Collection<NotificationGroup>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetStreamingEventsResults"/> class.
        /// </summary>
        internal GetStreamingEventsResults()
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.Notification);

            do
            {
                NotificationGroup notifications = new NotificationGroup();
                notifications.SubscriptionId = reader.ReadElementValue(XmlNamespace.Types, XmlElementNames.SubscriptionId);
                notifications.Events = new Collection<NotificationEvent>();

                lock (this)
                {
                    this.events.Add(notifications);
                }

                do
                {
                    reader.Read();

                    if (reader.IsStartElement())
                    {
                        string eventElementName = reader.LocalName;
                        EventType eventType;

                        if (GetEventsResults.XmlElementNameToEventTypeMap.TryGetValue(eventElementName, out eventType))
                        {
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
                                    eventType, 
                                    notifications);
                            }
                        }
                        else
                        {
                            reader.SkipCurrentElement();
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Notification));

                reader.Read();
            }
            while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.Notifications));
        }

        /// <summary>
        /// Loads a notification event from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="eventElementName">Name of the event XML element.</param>
        /// <param name="eventType">Type of the event.</param>
        /// <param name="notifications">Collection of notifications</param>
        private void LoadNotificationEventFromXml(
            EwsServiceXmlReader reader,
            string eventElementName,
            EventType eventType,
            NotificationGroup notifications)
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
            notifications.Events.Add(notificationEvent);
        }

        /// <summary>
        /// Gets the notification collection.
        /// </summary>
        /// <value>The notification collection.</value>
        internal Collection<NotificationGroup> Notifications
        {
            get { return this.events; }
        }
    }
}
