// ---------------------------------------------------------------------------
// <copyright file="NotificationEventArgs.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NotificationEventArgs class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Provides data to a StreamingSubscriptionConnection's OnNotificationEvent event.
    /// </summary>
    public class NotificationEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationEventArgs"/> class.
        /// </summary>
        /// <param name="subscription">The subscription for which notifications have been received.</param>
        /// <param name="events">The events that were received.</param>
        internal NotificationEventArgs(
            StreamingSubscription subscription,
            IEnumerable<NotificationEvent> events)
        {
            this.Subscription = subscription;
            this.Events = events;
        }

        /// <summary>
        /// Gets the subscription for which notifications have been received.
        /// </summary>
        public StreamingSubscription Subscription
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the events that were received.
        /// </summary>
        public IEnumerable<NotificationEvent> Events
        {
            get;
            internal set;
        }
    }
}