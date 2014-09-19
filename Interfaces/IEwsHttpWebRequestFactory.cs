//-----------------------------------------------------------------------
// <copyright file="IEwsHttpWebRequestFactory.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>Defines the IEwsHttpWebRequestFactory interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Net;

    /// <summary>
    /// Defines a factory interface for creating IEwsHttpWebRequest and IEwsHttpWebResponse instances.
    /// </summary>
    internal interface IEwsHttpWebRequestFactory
    {
        /// <summary>
        /// Create a new instance of class that implements the <see cref="IEwsHttpWebRequest"/> interface.
        /// </summary>
        /// <param name="uri">The URI.</param>
        /// <returns>
        /// An object that implements the <see cref="IEwsHttpWebRequest"/> interface.
        /// </returns>
        IEwsHttpWebRequest CreateRequest(Uri uri);

        /// <summary>
        /// Creates the exception response.
        /// </summary>
        /// <param name="exception">The exception.</param>
        /// <returns></returns>
        IEwsHttpWebResponse CreateExceptionResponse(WebException exception);
    }
}
