#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

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
