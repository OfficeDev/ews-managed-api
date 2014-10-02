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
// <summary>Defines the ClientCertificateCredentials class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Net;
    using System.Security.Cryptography.X509Certificates;

    /// <summary>
    /// ClientCertificateCredentials wraps an instance of X509CertificateCollection used for client certification-based authentication.
    /// </summary>
    public sealed class ClientCertificateCredentials : ExchangeCredentials
    {
        /// <summary>
        /// Collection of client certificates.
        /// </summary>
        private X509CertificateCollection clientCertificates;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClientCertificateCredentials"/> class.
        /// </summary>
        /// <param name="clientCertificates">The client certificates.</param>
        public ClientCertificateCredentials(X509CertificateCollection clientCertificates)
        {
            EwsUtilities.ValidateParam(clientCertificates, "clientCertificates");

            this.clientCertificates = clientCertificates;
        }

        /// <summary>
        /// This method is called to apply credentials to a service request before the request is made.
        /// </summary>
        /// <param name="request">The request.</param>
        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            request.ClientCertificates = this.ClientCertificates;
        }

        /// <summary>
        /// Gets the client certificates collection.
        /// </summary>
        public X509CertificateCollection ClientCertificates
        {
            get { return this.clientCertificates; }
        }
    }
}
