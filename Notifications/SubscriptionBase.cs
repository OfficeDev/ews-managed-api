// ---------------------------------------------------------------------------
// <copyright file="SubscriptionBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SubscriptionBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents the base class for event subscriptions.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class SubscriptionBase
    {
        private ExchangeService service;
        private string id;
        private string watermark;

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscriptionBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SubscriptionBase(ExchangeService service)
        {
            EwsUtilities.ValidateParam(service, "service");

            this.service = service;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscriptionBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="id">The id.</param>
        internal SubscriptionBase(ExchangeService service, string id)
            : this(service)
        {
            EwsUtilities.ValidateParam(id, "id");

            this.id = id;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscriptionBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="id">The id.</param>
        /// <param name="watermark">The watermark.</param>
        internal SubscriptionBase(
            ExchangeService service,
            string id,
            string watermark)
            : this(service, id)
        {
            this.watermark = watermark;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal virtual void LoadFromXml(EwsServiceXmlReader reader)
        {
            this.id = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.SubscriptionId);
            
            if (this.UsesWatermark)
            {
                this.watermark = reader.ReadElementValue(XmlNamespace.Messages, XmlElementNames.Watermark);
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonResponse">The json response.</param>
        /// <param name="service">The service.</param>
        internal virtual void LoadFromJson(JsonObject jsonResponse, ExchangeService service)
        {
            this.id = jsonResponse.ReadAsString(XmlElementNames.SubscriptionId);

            if (this.UsesWatermark)
            {
                this.watermark = jsonResponse.ReadAsString(XmlElementNames.Watermark);
            }
        }

        /// <summary>
        /// Gets the session.
        /// </summary>
        /// <value>The session.</value>
        internal ExchangeService Service
        {
            get { return this.service; }
        }

        /// <summary>
        /// Gets the Id of the subscription.
        /// </summary>
        public string Id
        {
            get { return this.id; }
            internal set { this.id = value; }
        }

        /// <summary>
        /// Gets the latest watermark of the subscription. Watermark is always null for streaming subscriptions.
        /// </summary>
        public string Watermark
        {
            get { return this.watermark; }
            internal set { this.watermark = value; }
        }

        /// <summary>
        /// Gets whether or not this subscription uses watermarks.
        /// </summary>
        protected virtual bool UsesWatermark
        {
            get { return true; }
        }
    }
}
