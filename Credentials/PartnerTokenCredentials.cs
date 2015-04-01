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
    using System.Globalization;
    using System.IO;
    using System.Net;
    using System.Security.Cryptography;
    using System.Security.Cryptography.Xml;
    using System.Xml;

    /// <summary>
    /// PartnerTokenCredentials can be used to send EWS or autodiscover requests to the managed tenant.
    /// </summary>
    internal sealed class PartnerTokenCredentials : WSSecurityBasedCredentials
    {
        private const string WsSecuritySymmetricKeyPathSuffix = "/wssecurity/symmetrickey";

        private readonly KeyInfoNode keyInfoNode;

        /// <summary>
        /// Initializes a new instance of the <see cref="PartnerTokenCredentials"/> class.
        /// </summary>
        /// <param name="securityToken">The token.</param>
        /// <param name="securityTokenReference">The token reference.</param>
        internal PartnerTokenCredentials(string securityToken, string securityTokenReference)
            : base(securityToken, true /* addTimestamp */)
        {
            EwsUtilities.ValidateParam(securityToken, "securityToken");
            EwsUtilities.ValidateParam(securityTokenReference, "securityTokenReference");

            SafeXmlDocument doc = new SafeXmlDocument();
            doc.PreserveWhitespace = true;
            doc.LoadXml(securityTokenReference);
            this.keyInfoNode = new KeyInfoNode(doc.DocumentElement);
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
            return new Uri(GetUriWithoutSuffix(url) + PartnerTokenCredentials.WsSecuritySymmetricKeyPathSuffix);
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

            //signedXml.AddReference("/soap:Envelope/soap:Header/t:ExchangeImpersonation");
            signedXml.AddReference("/soap:Envelope/soap:Header/wsse:Security/wsu:Timestamp");

            signedXml.KeyInfo.AddClause(this.keyInfoNode);
            using (var hashedAlgorithm = new HMACSHA1(ExchangeServiceBase.SessionKey))
            {
                signedXml.ComputeSignature(hashedAlgorithm);
            }

            XmlElement signature = signedXml.GetXml();

            XmlNode wssecurityNode = document.SelectSingleNode(
                "/soap:Envelope/soap:Header/wsse:Security",
                WSSecurityBasedCredentials.NamespaceManager);

            wssecurityNode.AppendChild(signature);

            memoryStream.Position = 0;
            document.Save(memoryStream);
        }
    }
}