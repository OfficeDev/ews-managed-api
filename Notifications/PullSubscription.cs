// ---------------------------------------------------------------------------
// <copyright file="PullSubscription.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PullSubscription class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a pull subscription.
    /// </summary>
    public sealed class PullSubscription : SubscriptionBase
    {
        private bool? moreEventsAvailable;

        /// <summary>
        /// Initializes a new instance of the <see cref="PullSubscription"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal PullSubscription(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Obtains a collection of events that occurred on the subscribed folders since the point
        /// in time defined by the Watermark property. When GetEvents succeeds, Watermark is updated.
        /// </summary>
        /// <returns>Returns a collection of events that occurred since the last watermark.</returns>
        public GetEventsResults GetEvents()
        {
            GetEventsResults results = this.Service.GetEvents(this.Id, this.Watermark);

            this.Watermark = results.NewWatermark;
            this.moreEventsAvailable = results.MoreEventsAvailable;

            return results;
        }

        /// <summary>
        /// Begins an asynchronous request to obtain a collection of events that occurred on the subscribed 
        /// folders since the point in time defined by the Watermark property.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginGetEvents(AsyncCallback callback, object state)
        {
            return this.Service.BeginGetEvents(callback, state, this.Id, this.Watermark);
        }

        /// <summary>
        /// Ends an asynchronous request to obtain a collection of events that occurred on the subscribed 
        /// folders since the point in time defined by the Watermark property. When EndGetEvents succeeds, Watermark is updated.
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>Returns a collection of events that occurred since the last watermark.</returns>
        public GetEventsResults EndGetEvents(IAsyncResult asyncResult)
        {
            GetEventsResults results = this.Service.EndGetEvents(asyncResult);

            this.Watermark = results.NewWatermark;
            this.moreEventsAvailable = results.MoreEventsAvailable;

            return results;
        }

        /// <summary>
        /// Unsubscribes from the pull subscription.
        /// </summary>
        public void Unsubscribe()
        {
            this.Service.Unsubscribe(this.Id);
        }

        /// <summary>
        /// Begins an asynchronous request to unsubscribe from the pull subscription. 
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        public IAsyncResult BeginUnsubscribe(AsyncCallback callback, object state)
        {
            return this.Service.BeginUnsubscribe(callback, state, this.Id);
        }

        /// <summary>
        /// Ends an asynchronous request to unsubscribe from the pull subscription. 
        /// </summary>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        public void EndUnsubscribe(IAsyncResult asyncResult)
        {
            this.Service.EndUnsubscribe(asyncResult);
        }

        /// <summary>
        /// Gets a value indicating whether more events are available on the server.
        /// MoreEventsAvailable is undefined (null) until GetEvents is called.
        /// </summary>
        public bool? MoreEventsAvailable
        {
            get { return this.moreEventsAvailable; }
        }
    }
}
