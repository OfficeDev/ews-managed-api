// ---------------------------------------------------------------------------
// <copyright file="ClientAccessTokenRequest.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ClientAccessTokenRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a client token access request
    /// </summary>
    public class ClientAccessTokenRequest : ComplexProperty
    {
        private readonly string id;
        private readonly ClientAccessTokenType tokenType;
        private readonly string scope;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClientAccessTokenRequest"/> class.
        /// </summary>
        /// <param name="id">id</param>
        /// <param name="tokenType">The tokenType.</param>
        public ClientAccessTokenRequest(string id, ClientAccessTokenType tokenType): this(id, tokenType, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ClientAccessTokenRequest"/> class.
        /// </summary>
        /// <param name="id">id</param>
        /// <param name="tokenType">The tokenType.</param>
        /// <param name="scope">The scope.</param>
        public ClientAccessTokenRequest(string id, ClientAccessTokenType tokenType, string scope)
        {
            this.id = id;
            this.tokenType = tokenType;
            this.scope = scope;
        }

        /// <summary>
        /// Gets the App Id.
        /// </summary>
        public string Id
        {
            get { return this.id; }
        }

        /// <summary>
        /// Gets token type.
        /// </summary>
        public ClientAccessTokenType TokenType
        {
            get { return this.tokenType; }
        }

        /// <summary>
        /// Gets the token scope.
        /// </summary>
        public string Scope
        {
            get { return this.scope; }
        }
    }
}
