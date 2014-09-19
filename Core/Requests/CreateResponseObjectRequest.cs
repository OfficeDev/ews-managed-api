// ---------------------------------------------------------------------------
// <copyright file="CreateResponseObjectRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CreateResponseObjectRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a CreateItem request for a response object.
    /// </summary>
    internal sealed class CreateResponseObjectRequest : CreateItemRequestBase<ServiceObject, CreateResponseObjectResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CreateResponseObjectRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal CreateResponseObjectRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service object.</returns>
        internal override CreateResponseObjectResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new CreateResponseObjectResponse();
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
