//-----------------------------------------------------------------------
// <copyright file="IEwsHttpWebResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>Defines the IEwsHttpWebResponse interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.Net;

    /// <summary>
    /// Interface representing HTTP web response.
    /// </summary>
    internal interface IEwsHttpWebResponse : IDisposable
    {
        /// <summary>
        /// Closes the response stream.
        /// </summary>
        void Close();

        /// <summary>
        /// Gets the stream that is used to read the body of the response from the server.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.IO.Stream"/> containing the body of the response.
        /// </returns>
        Stream GetResponseStream();

        /// <summary>
        /// Gets the method that is used to encode the body of the response.
        /// </summary>
        /// <returns>A string that describes the method that is used to encode the body of the response.</returns>
        string ContentEncoding { get; }

        /// <summary>
        /// Gets the content type of the response.
        /// </summary>
        /// <returns>A string that contains the content type of the response.</returns>
        string ContentType { get; }

        /// <summary>
        /// Gets the headers that are associated with this response from the server.
        /// </summary>
        /// <returns>A <see cref="T:System.Net.WebHeaderCollection"/> that contains the header information returned with the response.</returns>
        WebHeaderCollection Headers { get; }

        /// <summary>
        /// Gets the URI of the Internet resource that responded to the request.
        /// </summary>
        /// <returns>A <see cref="T:System.Uri"/> that contains the URI of the Internet resource that responded to the request.</returns>
        Uri ResponseUri { get; }

        /// <summary>
        /// Gets the status of the response.
        /// </summary>
        /// <returns>One of the System.Net.HttpStatusCode values.</returns>
        HttpStatusCode StatusCode { get; }

        /// <summary>
        /// Gets the status description returned with the response.
        /// </summary>
        /// <returns>A string that describes the status of the response.</returns>
        string StatusDescription { get; }

        /// <summary>
        /// Gets the version of the HTTP protocol that is used in the response.
        /// </summary>
        /// <returns>System.Version that contains the HTTP protocol version of the response.</returns>
        Version ProtocolVersion { get; }
    }
}
