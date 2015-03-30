/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Net;
    using System.Xml;

    /// <summary>
    /// WebCredentials wraps an instance of ICredentials used for password-based authentication schemes such as basic, digest, NTLM, and Kerberos authentication.
    /// </summary>
    public sealed class WebCredentials : ExchangeCredentials
    {
        private ICredentials credentials;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebCredentials"/> class to use
        /// the default network credentials.
        /// </summary>
        public WebCredentials()
            : this(CredentialCache.DefaultNetworkCredentials)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebCredentials"/> class using
        /// specified credentials.
        /// </summary>
        /// <param name="credentials">Credentials to use.</param>
        public WebCredentials(ICredentials credentials)
        {
            EwsUtilities.ValidateParam(credentials, "credentials");

            this.credentials = credentials;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebCredentials"/> class.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <param name="password">The password.</param>
        public WebCredentials(string username, string password)
            : this(new NetworkCredential(username, password))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebCredentials"/> class.
        /// </summary>
        /// <param name="username">Account username.</param>
        /// <param name="password">Account password.</param>
        /// <param name="domain">Account domain.</param>
        public WebCredentials(
            string username,
            string password,
            string domain)
            : this(new NetworkCredential(
                username, 
                password,
                domain))
        {
        }

        /// <summary>
        /// Applies NetworkCredential associated with this instance to a service request.
        /// </summary>
        /// <param name="request">The request.</param>
        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            request.Credentials = this.credentials;
        }

        /// <summary>
        /// Gets the Credentials from this instance.
        /// </summary>
        /// <value>The credentials.</value>
        public ICredentials Credentials
        {
            get { return this.credentials; }
        }

        /// <summary>
        /// Adjusts the URL endpoint based on the credentials. 
        /// For WebCredentials, the end user is responsible for setting the url.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns>The unchanged URL.</returns>
        internal override Uri AdjustUrl(Uri url)
        {
            return url;
        }
    }
}