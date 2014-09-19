// ---------------------------------------------------------------------------
// <copyright file="GetAppManifestsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetAppManifestsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents the response to a GetAppManifests operation.
    /// </summary>
    internal sealed class GetAppManifestsResponse : ServiceResponse
    {
        /// <summary>
        /// List of manifests returned in the response.
        /// </summary>
        private Collection<XmlDocument> manifests = new Collection<XmlDocument>();

        /// <summary>
        /// List of extensions returned in the response.
        /// </summary>
        private Collection<ClientApp> apps = new Collection<ClientApp>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetAppManifestsResponse"/> class.
        /// </summary>
        internal GetAppManifestsResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets all manifests returned
        /// </summary>
        /// <remarks>Provided for backwards compatibility with Exchange 2013.</remarks>
        public Collection<XmlDocument> Manifests
        {
            get { return this.manifests; }
        }

        /// <summary>
        /// Gets all apps returned.
        /// </summary>
        /// <remarks>Introduced for Exchange 2013 Sp1 to return additional metadata.</remarks>
        public Collection<ClientApp> Apps
        {
            get { return this.apps; }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.Manifests.Clear();
            base.ReadElementsFromXml(reader);

            reader.Read(XmlNodeType.Element);

            // We can have a response from Exchange 2013 (first time this API was introduced)
            // or the newer response, starting in Exchange 2013 SP1, (X-EWS-TargetVersion: 2.5 or above) 
            bool exchange2013Response;
            if (XmlElementNames.Manifests.Equals(reader.LocalName))
            {
                exchange2013Response = true;
            }
            else if (XmlElementNames.Apps.Equals(reader.LocalName))
            {
                exchange2013Response = false;
            }
            else
            {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.UnexpectedElement,
                        EwsUtilities.GetNamespacePrefix(XmlNamespace.Messages),
                        XmlElementNames.Manifests,
                        XmlNodeType.Element,
                        reader.LocalName,
                        reader.NodeType));
            }

            if (!reader.IsEmptyElement)
            {
                // Because we don't have an element for count of returned object,
                // we have to test the element to determine if it is StartElement of return object or EndElement
                reader.Read();

                if (exchange2013Response)
                {
                    this.ReadFromExchange2013(reader);
                }
                else
                {
                    this.ReadFromExchange2013Sp1(reader);
                }
            }

            reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Messages, exchange2013Response ? XmlElementNames.Manifests : XmlElementNames.Apps);
        }

        /// <summary>
        /// Read the response from Exchange 2013.
        /// This method assumes that the reader is currently at the Manifests element.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadFromExchange2013(EwsServiceXmlReader reader)
        {
            ////<GetAppManifestsResponse ResponseClass="Success" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
            ////<ResponseCode>NoError</ResponseCode>
            ////<m:Manifests xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">   
            ////<m:Manifest>[base 64 encoded manifest]</m:Manifest>                              <--- reader should be at this node at the beginning of loop
            ////<m:Manifest>[base 64 encoded manifest]</m:Manifest>
            //// ....
            ////</m:Manifests>                                                                   <--- reader should be at this node at the end of the loop
            while (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.Manifest))
            {
                XmlDocument manifest = ClientApp.ReadToXmlDocument(reader);
                this.Manifests.Add(manifest);

                this.Apps.Add(new ClientApp() { Manifest = manifest });
            }
        }

        /// <summary>
        /// Read the response from Exchange 2013.
        /// This method assumes that the reader is currently at the Manifests element.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadFromExchange2013Sp1(EwsServiceXmlReader reader)
        {
            ////<GetAppManifestsResponse ResponseClass="Success" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
            ////  <ResponseCode>NoError</ResponseCode>
            ////  <m:Apps xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
            ////    <t:App xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">       <--- reader should be at this node at the beginning of the loop
            ////      <t:Metadata>
            ////        <t:EndNodeUrl>http://o15.officeredir.microsoft.com/r/rlidMktplcExchRedirect?app=outlook.exe&amp;ver=15&amp;clid=1033&amp;p1=15d0d766d0&amp;p2=4&amp;p3=0&amp;p4=WA&amp;p5=en-US\WA102996382&amp;Scope=2&amp;CallBackURL=https%3a%2f%2fexhv-4880%2fecp%2fExtension%2finstallFromURL.slab%3fexsvurl%3d1&amp;DeployId=EXHV-4680dom.extest.microsoft.com</t:EndNodeUrl>
            ////        <t:AppStatus>2.3</t:AppStatus>
            ////        <t:ActionUrl>http://o15.officeredir.microsoft.com/r/rlidMktplcExchRedirect?app=outlook.exe&amp;ver=15&amp;clid=1033&amp;p1=15d0d766d0&amp;p2=4&amp;p3=0&amp;p4=WA&amp;p5=en-US\WA102996382&amp;Scope=2&amp;CallBackURL=https%3a%2f%2fexhv-4880%2fecp%2fExtension%2finstallFromURL.slab%3fexsvurl%3d1&amp;DeployId=EXHV-4680dom.extest.microsoft.com</t:ActionUrl>
            ////      </t:Metadata>
            ////      <t:Manifest>[base 64 encoded manifest]</t:Manifest>
            ////    </t:App>
            ////    <t:App xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            ////      ....
            ////  <m:Apps>    <----- reader should be at this node at the end of the loop
            while (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.App))
            {
                ClientApp clientApp = new ClientApp();
                clientApp.LoadFromXml(reader, XmlElementNames.App);

                this.Apps.Add(clientApp);
                this.Manifests.Add(clientApp.Manifest);

                reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Types, XmlElementNames.App);
                reader.Read();
            }
        }
    }
}
