// ---------------------------------------------------------------------------
// <copyright file="WebClientUrlCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the WebClientUrlCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a user setting that is a collection of Exchange web client URLs.
    /// </summary>
    public sealed class WebClientUrlCollection
    {
        private List<WebClientUrl> urls;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebClientUrlCollection"/> class.
        /// </summary>
        internal WebClientUrlCollection()
        {
            this.urls = new List<WebClientUrl>();
        }

        /// <summary>
        /// Loads instance of WebClientUrlCollection from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal static WebClientUrlCollection LoadFromXml(EwsXmlReader reader)
        {
            WebClientUrlCollection instance = new WebClientUrlCollection();

            do
            {
                reader.Read();

                if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.WebClientUrl))
                {
                    instance.Urls.Add(WebClientUrl.LoadFromXml(reader));
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.WebClientUrls));

            return instance;
        }

        /// <summary>
        /// Gets the URLs.
        /// </summary>
        public List<WebClientUrl> Urls
        {
            get { return this.urls; }
        }
    }
}
