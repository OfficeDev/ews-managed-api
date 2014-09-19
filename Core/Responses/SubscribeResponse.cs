// ---------------------------------------------------------------------------
// <copyright file="SubscribeResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SubscribeResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the base response class to subscription creation operations.
    /// </summary>
    /// <typeparam name="TSubscription">Subscription type.</typeparam>
    internal sealed class SubscribeResponse<TSubscription> : ServiceResponse
        where TSubscription : SubscriptionBase
    {
        private TSubscription subscription;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscribeResponse&lt;TSubscription&gt;"/> class.
        /// </summary>
        /// <param name="subscription">The subscription.</param>
        internal SubscribeResponse(TSubscription subscription)
            : base()
        {
            EwsUtilities.Assert(
                subscription != null,
                "SubscribeResponse.ctor",
                "subscription is null");

            this.subscription = subscription;
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.subscription.LoadFromXml(reader);
        }

        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            base.ReadElementsFromJson(responseObject, service);

            this.subscription.LoadFromJson(responseObject, service);
        }

        /// <summary>
        /// Gets the subscription that was created.
        /// </summary>
        public TSubscription Subscription
        {
            get { return this.subscription; }
        }
    }
}
