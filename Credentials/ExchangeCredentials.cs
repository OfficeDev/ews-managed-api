// ---------------------------------------------------------------------------
// <copyright file="ExchangeCredentials.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ExchangeCredentials class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.Net;
    using System.Xml;

    /// <summary>
    /// Base class of Exchange credential types.
    /// </summary>
    public abstract class ExchangeCredentials
    {
        /// <summary>
        /// Performs an implicit conversion from <see cref="System.Net.NetworkCredential"/> to <see cref="Microsoft.Exchange.WebServices.Data.ExchangeCredentials"/>.
        /// This allows a NetworkCredential object to be implictly converted to an ExchangeCredential which is useful when setting
        /// credentials on an ExchangeService.
        /// </summary>
        /// <example>
        /// This operator allows you to type:
        /// <code>service.Credentials = new NetworkCredential("username","password");</code>
        /// instead of:
        /// <code>service.Credentials = new WebCredentials(new NetworkCredential("username","password"));</code>
        /// </example>
        /// <param name="credentials">The credentials.</param>
        /// <returns>The result of the conversion.</returns>
        public static implicit operator ExchangeCredentials(NetworkCredential credentials)
        {
            return new WebCredentials(credentials);
        }

        /// <summary>
        /// Performs an implicit conversion from <see cref="System.Net.CredentialCache"/> to <see cref="Microsoft.Exchange.WebServices.Data.ExchangeCredentials"/>.
        /// This allows a CredentialCache object to be implictly converted to an ExchangeCredential which is useful when setting
        /// credentials on an ExchangeService.
        /// </summary>
        /// <example>
        /// Using these credentials:
        /// <code>CredentialCache credentials = new CredentialCache();</code>
        /// <code>credentials.Add(new Uri("http://www.contoso.com/"),"Basic",new NetworkCredential(user,pwd));</code>
        /// <code>credentials.Add(new Uri("http://www.contoso.com/"),"Digest", new NetworkCredential(user,pwd,domain));</code>
        /// This operator allows you to type:
        /// <code>service.Credentials = credentials;</code>
        /// instead of:
        /// <code>service.Credentials = new WebCredentials(credentials);</code>
        /// </example>
        /// <param name="credentials">The credentials.</param>
        /// <returns>The result of the conversion.</returns>
        public static implicit operator ExchangeCredentials(CredentialCache credentials)
        {
            return new WebCredentials(credentials);
        }

        /// <summary>
        /// Return the url without suffix.
        /// </summary>
        /// <param name="url">The url</param>
        /// <returns>The absolute uri base.</returns>
        internal static string GetUriWithoutSuffix(Uri url)
        {
            string absoluteUri = url.AbsoluteUri;

            int index = absoluteUri.IndexOf(WSSecurityBasedCredentials.WsSecurityPathSuffix, StringComparison.OrdinalIgnoreCase);
            if (index != -1)
            {
                return absoluteUri.Substring(0, index);
            }

            return absoluteUri;
        }

        /// <summary>
        /// This method is called to pre-authenticate credentials before a service request is made.  
        /// </summary>
        internal virtual void PreAuthenticate()
        {
            // do nothing by default.
        }

        /// <summary>
        /// This method is called to apply credentials to a service request before the request is made.  
        /// </summary>
        /// <param name="request">The request.</param>
        internal virtual void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            // do nothing by default.
        }

        /// <summary>
        /// Emit any extra necessary namespace aliases for the SOAP:header block.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void EmitExtraSoapHeaderNamespaceAliases(XmlWriter writer)
        {
            // do nothing by default.
        }

        /// <summary>
        /// Serialize any extra necessary SOAP headers.
        /// This is used for authentication schemes that rely on WS-Security, or for endpoints requiring WS-Addressing.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="webMethodName">The Web method being called.</param>
        internal virtual void SerializeExtraSoapHeaders(XmlWriter writer, string webMethodName)
        {
            // do nothing by default.
        }

        /// <summary>
        /// Serialize SOAP headers used for authentication schemes that rely on WS-Security
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal virtual void SerializeWSSecurityHeaders(XmlWriter writer)
        {
            // do nothing by default.
        }

        /// <summary>
        /// Adjusts the URL endpoint based on the credentials.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns>Adjust URL.</returns>
        internal virtual Uri AdjustUrl(Uri url)
        {
            return new Uri(GetUriWithoutSuffix(url));
        }

        /// <summary>
        /// Gets the flag indicating whether any sign action need taken.
        /// </summary>
        internal virtual bool NeedSignature
        {
            get { return false; }
        }

        /// <summary>
        /// Add the signature element to the memory stream.
        /// </summary>
        /// <param name="memoryStream">The memory stream.</param>
        internal virtual void Sign(MemoryStream memoryStream)
        {
            throw new InvalidOperationException();
        }
    }
}
