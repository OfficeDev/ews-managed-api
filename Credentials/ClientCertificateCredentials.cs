// ---------------------------------------------------------------------------
// <copyright file="ClientCertificateCredentials.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
