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