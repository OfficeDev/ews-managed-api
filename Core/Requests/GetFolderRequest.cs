// ---------------------------------------------------------------------------
// <copyright file="GetFolderRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetFolderRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a GetFolder request.
    /// </summary>
    internal sealed class GetFolderRequest : GetFolderRequestBase<GetFolderResponse>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetFolderRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal GetFolderRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service, errorHandlingMode)
        {
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal override GetFolderResponse CreateServiceResponse(ExchangeService service, int responseIndex)
        {
            return new GetFolderResponse(this.FolderIds[responseIndex].GetFolder(), this.PropertySet);
        }
    }
}
