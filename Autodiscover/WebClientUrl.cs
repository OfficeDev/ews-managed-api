// ---------------------------------------------------------------------------
// <copyright file="WebClientUrl.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the WebClientUrl class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the URL of the Exchange web client.
    /// </summary>
    public sealed class WebClientUrl
    {
        private string authenticationMethods;
        private string url;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebClientUrl"/> class.
        /// </summary>
        private WebClientUrl()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WebClientUrl"/> class.
        /// </summary>
        /// <param name="authenticationMethods">The authentication methods.</param>
        /// <param name="url">The URL.</param>
        internal WebClientUrl(string authenticationMethods, string url)
        {
            this.authenticationMethods = authenticationMethods;
            this.url = url;
        }

        /// <summary>
        /// Loads WebClientUrl instance from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>WebClientUrl.</returns>
        internal static WebClientUrl LoadFromXml(EwsXmlReader reader)
        {
            WebClientUrl webClientUrl = new WebClientUrl();

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.AuthenticationMethods:
                            webClientUrl.AuthenticationMethods = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.Url:
                            webClientUrl.Url = reader.ReadElementValue<string>();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.WebClientUrl));

            return webClientUrl;
        }

        /// <summary>
        /// Gets the authentication methods.
        /// </summary>
        public string AuthenticationMethods
        {
            get { return this.authenticationMethods; }
            internal set { this.authenticationMethods = value; }
        }

        /// <summary>
        /// Gets the URL.
        /// </summary>
        public string Url
        {
            get { return this.url; }
            internal set { this.url = value; }
        }
    }
}
