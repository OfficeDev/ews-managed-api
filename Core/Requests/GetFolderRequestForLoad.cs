// ---------------------------------------------------------------------------
// <copyright file="GetFolderRequestForLoad.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetFolderRequestForLoad class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetFolder request specialized to return ServiceResponse.
    /// </summary>
    internal sealed class GetFolderRequestForLoad : GetFolderRequestBase<ServiceResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetFolderRequestForLoad"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal GetFolderRequestForLoad(ExchangeService service, ServiceErrorHandling errorHandlingMode)
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
            return new GetFolderResponse(this.FolderIds[responseIndex].GetFolder(), this.PropertySet);
        }
    }
}
