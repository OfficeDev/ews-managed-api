// ---------------------------------------------------------------------------
// <copyright file="OAuthCredentials.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Net;
    using System.Text.RegularExpressions;

    /// <summary>
    /// OAuthCredentials provides credentials for server-to-server authentication. The JSON web token is 
    /// defined at http://tools.ietf.org/id/draft-jones-json-web-token-03.txt. The token string is 
    /// base64url encoded (described in http://www.ietf.org/rfc/rfc4648.txt, section 5).
    /// 
    /// OAuthCredentials is supported for Exchange 2013 or above.
    /// </summary>
    public sealed class OAuthCredentials : ExchangeCredentials
    {
        private const string BearerAuthenticationType = "Bearer";

        private static readonly Regex validTokenPattern = new Regex(
            @"^[A-Za-z0-9-_]+\.[A-Za-z0-9-_]+\.[A-Za-z0-9-_]*$",
            RegexOptions.Compiled);

        private readonly string token;

        private readonly ICredentials credentials;

        /// <summary>
        /// Initializes a new instance of the <see cref="OAuthCredentials"/> class.
        /// </summary>
        /// <param name="token">The JSON web token string.</param>
        public OAuthCredentials(string token)
            : this(token, false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OAuthCredentials"/> class.
        /// </summary>
        /// <param name="token"></param>
        /// <param name="verbatim"></param>
        internal OAuthCredentials(string token, bool verbatim)
        {
            EwsUtilities.ValidateParam(token, "token");

            string rawToken;
            if (verbatim)
            {
                rawToken = token;
            }
            else
            {
                int whiteSpacePosition = token.IndexOf(' ');
                if (whiteSpacePosition == -1)
                {
                    rawToken = token;
                }
                else
                {
                    string authType = token.Substring(0, whiteSpacePosition);
                    if (string.Compare(authType, BearerAuthenticationType, StringComparison.OrdinalIgnoreCase) != 0)
                    {
                        throw new ArgumentException(Strings.InvalidAuthScheme);
                    }

                    rawToken = token.Substring(whiteSpacePosition + 1);
                }

                if (!validTokenPattern.IsMatch(rawToken))
                {
                    throw new ArgumentException(Strings.InvalidOAuthToken);
                }
            }

            this.token = BearerAuthenticationType + " " + rawToken;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OAuthCredentials"/> class using
        /// specified credentials.
        /// </summary>
        /// <param name="credentials">Credentials to use.</param>
        public OAuthCredentials(ICredentials credentials)
        {
            EwsUtilities.ValidateParam(credentials, "credentials");

            this.credentials = credentials;
        }

        /// <summary>
        /// Add the Authorization header to a service request.
        /// </summary>
        /// <param name="request">The request</param>
        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            base.PrepareWebRequest(request);

            if (this.token != null)
            {
                request.Headers.Remove(HttpRequestHeader.Authorization);
                request.Headers.Add(HttpRequestHeader.Authorization, this.token);
            }
            else
            {
                request.Credentials = this.credentials;
            }
        }
    }
}