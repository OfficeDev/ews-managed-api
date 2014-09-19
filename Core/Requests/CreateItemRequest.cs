// ---------------------------------------------------------------------------
// <copyright file="CreateItemRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateItemRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a CreateItem request.
    /// </summary>
    internal sealed class CreateItemRequest : CreateItemRequestBase<Item, ServiceResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CreateItemRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal CreateItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override ServiceResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new CreateItemResponse((Item)EwsUtilities.GetEnumeratedObjectAt(this.Items, responseIndex));
        }

        /// <summary>
        /// Validate request..
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            // Validate each item.
            foreach (Item item in this.Items)
            {
                item.Validate();
            }
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }
    }
}
