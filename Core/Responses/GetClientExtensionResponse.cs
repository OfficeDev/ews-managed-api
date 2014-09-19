// ---------------------------------------------------------------------------
// <copyright file="GetClientExtensionResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetClientExtensionResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents the response to a GetClientExtension operation.
    /// </summary>
    public sealed class GetClientExtensionResponse : ServiceResponse
    {
        private Collection<ClientExtension> clientExtension = new Collection<ClientExtension>();

        private string rawMasterTableXml;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetClientExtensionResponse"/> class.
        /// </summary>
        internal GetClientExtensionResponse()
            : base()
        {
        }

        /// <summary>
        /// Gets all ClientExtension returned
        /// </summary>
        public Collection<ClientExtension> ClientExtensions
        {
            get { return this.clientExtension; }
        }

        /// <summary>
        /// Gets org raw master table xml
        /// </summary>
        public string RawMasterTableXml
        {
            get { return this.rawMasterTableXml; }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            this.ClientExtensions.Clear();
            base.ReadElementsFromXml(reader);
            
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ClientExtensions);

            if (!reader.IsEmptyElement)
            {
                // Because we don't have an element for count of returned object,
                // we have to test the element to determine if it is StartElement of return object or EndElement
                reader.Read();

                while (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ClientExtension))
                {
                    ClientExtension clientExtension = new ClientExtension();
                    clientExtension.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ClientExtension);
                    this.ClientExtensions.Add(clientExtension);

                    reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Types, XmlElementNames.ClientExtension);
                    reader.Read();
                }

                reader.EnsureCurrentNodeIsEndElement(XmlNamespace.Messages, XmlElementNames.ClientExtensions);
            }

            reader.Read();
            if (reader.IsStartElement(XmlNamespace.Messages, XmlElementNames.ClientExtensionRawMasterTableXml))
            {
                this.rawMasterTableXml = reader.ReadElementValue();
            }
        }
    }
}
