// ---------------------------------------------------------------------------
// <copyright file="MoveCopyItemRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MoveCopyItemRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an abstract Move/Copy Item request.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class MoveCopyItemRequest<TResponse> : MoveCopyRequest<Item, TResponse>
        where TResponse : ServiceResponse
    {
        private ItemIdWrapperList itemIds = new ItemIdWrapperList();

        /// <summary>
        /// Validates request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateParam(this.ItemIds, "ItemIds");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MoveCopyItemRequest&lt;TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal MoveCopyItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Writes the ids and returnNewItemIds flag as XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteIdsToXml(EwsServiceXmlWriter writer)
        {
            this.ItemIds.WriteToXml(
                writer,
                XmlNamespace.Messages,
                XmlElementNames.ItemIds);

            if (this.ReturnNewItemIds.HasValue)
            {
                writer.WriteElementValue(
                    XmlNamespace.Messages,
                    XmlElementNames.ReturnNewItemIds,
                    this.ReturnNewItemIds.Value);
            }
        }

        /// <summary>
        /// Adds the ids to json.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="service">The service.</param>
        internal override void AddIdsToJson(JsonObject jsonObject, ExchangeService service)
        {
            jsonObject.Add(XmlElementNames.ItemIds, this.ItemIds.InternalToJson(service));

            if (this.ReturnNewItemIds.HasValue)
            {
                jsonObject.Add(
                    XmlElementNames.ReturnNewItemIds,
                    this.ReturnNewItemIds.Value);
            }
        }

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
        {
            return this.ItemIds.Count;
        }

        /// <summary>
        /// Gets the item ids.
        /// </summary>
        /// <value>The item ids.</value>
        internal ItemIdWrapperList ItemIds
        {
            get { return this.itemIds; }
        }

        /// <summary>
        /// Gets or sets flag indicating whether we require that the service return new item ids.
        /// </summary>
        internal bool? ReturnNewItemIds
        {
            get; set;
        }
    }
}
