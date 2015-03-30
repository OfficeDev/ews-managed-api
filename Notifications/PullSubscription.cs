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