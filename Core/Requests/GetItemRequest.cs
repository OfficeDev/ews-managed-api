// ---------------------------------------------------------------------------
// <copyright file="GetItemRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetItemRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a GetItem request.
    /// </summary>
    internal sealed class GetItemRequest : GetItemRequestBase<GetItemResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetItemRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal GetItemRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override GetItemResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new GetItemResponse(this.ItemIds[responseIndex], this.PropertySet);
        }
    }
}
