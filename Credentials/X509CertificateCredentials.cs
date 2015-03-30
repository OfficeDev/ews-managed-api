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
    using System.IO;
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    using System.Security.Cryptography.Xml;
    using System.Xml;

    /// <summary>
    /// X509CertificateCredentials wraps an instance of X509Certificate2, it can be used for WS-Security/X509 certificate-based authentication.
    /// </summary>
    public sealed class X509CertificateCredentials : WSSecurityBasedCredentials
    {
        private const string BinarySecurityTokenFormat =
            "<wsse:BinarySecurityToken " +
            "EncodingType=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary\" " +
            "ValueType=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3\" " +
            "wsu:Id=\"{0}\">" +
            "{1}" +
            "</wsse:BinarySecurityToken>";

        private const string KeyInfoClauseFormat =
            "<wsse:SecurityTokenReference xmlns:wsse=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd\" >" +
            "<wsse:Reference URI=\"#{0}\" ValueType=\"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3\" />" +
            "</wsse:SecurityTokenReference>";

        private const string WsSecurityX509CertPathSuffix = "/wssecurity/x509cert";

        private readonly X509Certificate2 certificate;

        private readonly KeyInfoClause keyInfoClause;

        /// <summary>
        /// Initializes a new instance of the <see cref="X509CertificateCredentials"/> class.
        /// </summary>
        /// <remarks>The X509Certificate2 argument should have private key in order to sign the message.</remarks>
        /// <param name="certificate">The X509Certificate2 object.</param>
        public X509CertificateCredentials(X509Certificate2 certificate)
            : base(null, true)
        {
            EwsUtilities.ValidateParam(certificate, "certificate");

            if (!certificate.HasPrivateKey)
            {
                throw new ServiceValidationException(Strings.CertificateHasNoPrivateKey);
            }

            this.certificate = certificate;

            string certId = WSSecurityUtilityIdSignedXml.GetUniqueId();

            this.SecurityToken = string.Format(
                X509CertificateCredentials.BinarySecurityTokenFormat,
                certId,
                Convert.ToBase64String(this.certificate.GetRawCertData()));

            SafeXmlDocument doc = new SafeXmlDocument();
            doc.PreserveWhitespace = true;
            doc.LoadXml(string.Format(X509CertificateCredentials.KeyInfoClauseFormat, certId));
            this.keyInfoClause = new KeyInfoNode(doc.DocumentElement);
        }

        /// <summary>
        /// This method is called to apply credentials to a service request before the request is made.
        /// </summary>
        /// <param name="request">The request.</param>
        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            this.EwsUrl = request.RequestUri;
        }

        /// <summary>
        /// Adjusts the URL based on the credentials.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns>Adjust URL.</returns>
        internal override Uri AdjustUrl(Uri url)
        {
            return new Uri(GetUriWithoutSuffix(url) + X509CertificateCredentials.WsSecurityX509CertPathSuffix);
        }

        /// <summary>
        /// Gets the flag indicating whether any sign action need taken.
        /// </summary>
        internal override bool NeedSignature
        {
            get { return true; }
        }

        /// <summary>
        /// Add the signature element to the memory stream.
        /// </summary>
        /// <param name="memoryStream">The memory stream.</param>
        internal override void Sign(MemoryStream memoryStream)
        {
            memoryStream.Position = 0;

            SafeXmlDocument document = new SafeXmlDocument();
            document.PreserveWhitespace = true;
            document.Load(memoryStream);

            WSSecurityUtilityIdSignedXml signedXml = new WSSecurityUtilityIdSignedXml(document);
            signedXml.SignedInfo.CanonicalizationMethod = SignedXml.XmlDsigExcC14NTransformUrl;

            signedXml.SigningKey = this.certificate.PrivateKey;
            signedXml.AddReference("/soap:Envelope/soap:Header/wsa:To");
            signedXml.AddReference("/soap:Envelope/soap:Header/wsse:Security/wsu:Timestamp");

            signedXml.KeyInfo.AddClause(this.keyInfoClause);
            signedXml.ComputeSignature();
            XmlElement signature = signedXml.GetXml();

            XmlNode wssecurityNode = document.SelectSingleNode(
                "/soap:Envelope/soap:Header/wsse:Security",
                WSSecurityBasedCredentials.NamespaceManager);

            wssecurityNode.AppendChild(signature);

            memoryStream.Position = 0;
            document.Save(memoryStream);
        }

        /// <summary>
        /// Gets the credentials string presentation.
        /// </summary>
        /// <returns>The string.</returns>
        public override string ToString()
        {
            return string.Format(
                "X509:<I>={0},<S>={1}",
                this.certificate.Issuer,
                this.certificate.Subject);
        }
    }
}