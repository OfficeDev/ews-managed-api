// ---------------------------------------------------------------------------
// <copyright file="TokenCredentials.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TokenCredentials class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Net;
    using System.Xml;

    /// <summary>
    /// TokenCredentials provides credentials if you already have a token.
    /// </summary>
    public sealed class TokenCredentials : WSSecurityBasedCredentials
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TokenCredentials"/> class.
        /// </summary>
        /// <param name="securityToken">The token.</param>
        public TokenCredentials(string securityToken) 
            : base(securityToken)
        {
            EwsUtilities.ValidateParam(securityToken, "securityToken");
        }
        
        /// <summary>
        /// This method is called to apply credentials to a service request before the request is made.
        /// </summary>
        /// <param name="request">The request.</param>
        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            this.EwsUrl = request.RequestUri;
        }
    }
}
