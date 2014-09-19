// ---------------------------------------------------------------------------
// <copyright file="StreamingSubscription.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the StreamingSubscription class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a streaming subscription.
    /// </summary>
    public sealed class StreamingSubscription : SubscriptionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StreamingSubscription"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal StreamingSubscription(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Unsubscribes from the streaming subscription.
        /// </summary>
        public void Unsubscribe()
        {
            this.Service.Unsubscribe(this.Id);
        }

        /// <summary>
        /// Begins an asynchronous request to unsubscribe from the streaming subscription. 
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginUnsubscribe(AsyncCallback callback, object state)
        {
            return this.Service.BeginUnsubscribe(callback, state, this.Id);
        }

        /// <summary>
        /// Ends an asynchronous request to unsubscribe from the streaming subscription. 
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        public void EndUnsubscribe(IAsyncResult asyncResult)
        {
            this.Service.EndUnsubscribe(asyncResult);
        }

        /// <summary>
        /// Gets the service used to create this subscription.
        /// </summary>
        public new ExchangeService Service
        {
            get
            {
                return base.Service;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this subscription uses watermarks.
        /// </summary>
        protected override bool UsesWatermark
        {
            get
            {
                return false;
            }
        }
    }
}
