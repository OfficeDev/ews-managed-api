// ---------------------------------------------------------------------------
// <copyright file="PushSubscription.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PushSubscription class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a push subscriptions.
    /// </summary>
    public sealed class PushSubscription : SubscriptionBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PushSubscription"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal PushSubscription(ExchangeService service)
            : base(service)
        {
        }
    }
}
