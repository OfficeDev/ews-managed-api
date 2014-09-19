// ---------------------------------------------------------------------------
// <copyright file="SubscribeToPullNotificationsRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SubscribeToPullNotificationsRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a "pull" Subscribe request.
    /// </summary>
    internal class SubscribeToPullNotificationsRequest : SubscribeRequest<PullSubscription>
    {
        private int timeout = 30;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscribeToPullNotificationsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SubscribeToPullNotificationsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            if ((this.Timeout < 1) || (this.Timeout > 1440))
            {
                throw new ArgumentException(string.Format(Strings.InvalidTimeoutValue, this.Timeout));
            }
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override SubscribeResponse<PullSubscription> CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new SubscribeResponse<PullSubscription>(new PullSubscription(service));
        }

        /// <summary>
        /// Gets the name of the subscription XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetSubscriptionXmlElementName()
        {
            return XmlElementNames.PullSubscriptionRequest;
        }

        /// <summary>
        /// Internal method to write XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Timeout,
                this.Timeout);
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonSubscribeRequest">The json subscribe request.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonSubscribeRequest, ExchangeService service)
        {
            jsonSubscribeRequest.Add(XmlElementNames.Timeout, this.Timeout);
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Gets or sets the timeout.
        /// </summary>
        /// <value>The timeout.</value>
        public int Timeout
        {
            get { return this.timeout; }
            set { this.timeout = value; }
        }
    }
}
