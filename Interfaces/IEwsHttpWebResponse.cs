#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion
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
