//-----------------------------------------------------------------------
// <copyright file="EwsHttpWebRequestFactory.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>Defines the EwsHttpWebRequestFactory class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Net;

    /// <summary>
    /// Represents an implementation of IEwsHttpWebRequestFactory using EwsHttpWebRequest.
    /// </summary>
    internal class EwsHttpWebRequestFactory : IEwsHttpWebRequestFactory
    {
        #region IEwsHttpWebRequestFactory Members

        /// <summary>
        /// Create a new instance of <see cref="EwsHttpWebRequest"/>.
        /// </summary>
        /// <param name="uri">The service URI.</param>
        /// <returns>An instance of <see cref="IEwsHttpWebRequest"/>./// </returns>
        IEwsHttpWebRequest IEwsHttpWebRequestFactory.CreateRequest(Uri uri)
        {
            return new EwsHttpWebRequest(uri);
        }

        /// <summary>
        /// Creates response from a WebException.
        /// </summary>
        /// <param name="exception">The exception.</param>
        /// <returns>Instance of IEwsHttpWebResponse.</returns>
        IEwsHttpWebResponse IEwsHttpWebRequestFactory.CreateExceptionResponse(WebException exception)
        {
            EwsUtilities.ValidateParam(exception, "exception");

            if (exception.Response == null)
            {
                throw new InvalidOperationException("The exception does not contain response.");
            }

            return new EwsHttpWebResponse(exception.Response as HttpWebResponse);
        }
        #endregion
    }
}
