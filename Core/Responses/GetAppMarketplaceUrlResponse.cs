// ---------------------------------------------------------------------------
// <copyright file="GetAppMarketplaceUrlResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetAppMarketplaceUrlResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a GetAppMarketplaceUrl operation
    /// </summary>
    internal sealed class GetAppMarketplaceUrlResponse : ServiceResponse
    {
        private string appMarketplaceUrl;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetAppMarketplaceUrlResponse"/> class.
        /// </summary>
        internal GetAppMarketplaceUrlResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);
            this.appMarketplaceUrl = reader.ReadElementValue<string>(XmlNamespace.NotSpecified, XmlElementNames.AppMarketplaceUrl);
        }

        /// <summary>
        /// App Marketplace Url
        /// </summary>
        public string AppMarketplaceUrl
        {
            get { return this.appMarketplaceUrl; }
        }
    }
}
