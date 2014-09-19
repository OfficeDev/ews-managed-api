// ---------------------------------------------------------------------------
// <copyright file="PartnerTokenCredentials.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
