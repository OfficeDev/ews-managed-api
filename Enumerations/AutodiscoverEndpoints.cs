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
// <summary>Defines the AutodiscoverEndpoints enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the types of Autodiscover endpoints that are available.
    /// </summary>
    [Flags]
    internal enum AutodiscoverEndpoints
    {
        /// <summary>
        /// No endpoints available.
        /// </summary>
        None = 0,

        /// <summary>
        /// The "legacy" Autodiscover endpoint.
        /// </summary>
        Legacy = 1,

        /// <summary>
        /// The SOAP endpoint.
        /// </summary>
        Soap = 2,

        /// <summary>
        /// The WS-Security endpoint.
        /// </summary>
        WsSecurity = 4,

        /// <summary>
        /// The WS-Security/SymmetricKey endpoint.
        /// </summary>
        WSSecuritySymmetricKey = 8,

        /// <summary>
        /// The WS-Security/X509Cert endpoint.
        /// </summary>
        WSSecurityX509Cert = 16,

        /// <summary>
        /// The OAuth endpoint
        /// </summary>
        OAuth = 32,
    }
}
