// ---------------------------------------------------------------------------
// <copyright file="SubscriptionErrorEventArgs.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SubscriptionErrorEventArgs class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// 
    /// Provides data to a StreamingSubscriptionConnection's OnSubscriptionError and OnDisconnect events.
    /// </summary>
    public class SubscriptionErrorEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SubscriptionErrorEventArgs"/> class.
        /// </summary>
        /// <param name="subscription">The subscription for which an error occurred. If subscription is null, the error applies to the entire connection.</param>
        /// <param name="exception">The exception representing the error. If exception is null, the connection was cleanly closed by the server.</param>
        internal SubscriptionErrorEventArgs(
            StreamingSubscription subscription,
            Exception exception)
        {
            this.Subscription = subscription;
            this.Exception = exception;
        }

        /// <summary>
        /// Gets the subscription for which an error occurred. If Subscription is null, the error applies to the entire connection.
        /// </summary>
        public StreamingSubscription Subscription
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the exception representing the error. If Exception is null, the connection was cleanly closed by the server.
        /// </summary>
        public Exception Exception
        {
            get;
            internal set;
        }
    }
}