#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the ClientExtension class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// Represents a ClientExtension object.
    /// </summary>
    public sealed class ClientExtension : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ClientExtension"/> class.
        /// </summary>
        internal ClientExtension()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="ClientExtension"/> class.
        /// </summary>
        /// <param name="type">Extension type</param>
        /// <param name="scope">Extension install scope</param>
        /// <param name="manifestStream">Manifest stream, can be null</param>
        /// <param name="marketplaceAssetID">The asset ID for Office Marketplace</param>
        /// <param name="marketplaceContentMarket">The content market for Office Marketplace</param>
        /// <param name="isAvailable">Whether extension is available</param>
        /// <param name="isMandatory">Whether extension is mandatory</param>
        /// <param name="isEnabledByDefault">Whether extension is enabled by default</param>
        /// <param name="providedTo">Who the extension is provided for (e.g. "entire org" or "specific users")</param>
        /// <param name="specificUsers">List of users extension is provided for, can be null</param>
        /// <param name="appStatus">App status</param>
        /// <param name="etoken">Etoken</param>
        public ClientExtension(
            ExtensionType type,
            ExtensionInstallScope scope,
            Stream manifestStream,
            string marketplaceAssetID,
            string marketplaceContentMarket,
            bool isAvailable,
            bool isMandatory,
            bool isEnabledByDefault,
            ClientExtensionProvidedTo providedTo,
            StringList specificUsers,
            string appStatus,
            string etoken)
                : this()
        {
            this.Type = type;
            this.Scope = scope;
            this.ManifestStream = manifestStream;
            this.MarketplaceAssetID = marketplaceAssetID;
            this.MarketplaceContentMarket = marketplaceContentMarket;
            this.IsAvailable = isAvailable;
            this.IsMandatory = isMandatory;
            this.IsEnabledByDefault = isEnabledByDefault;
            this.ProvidedTo = providedTo;
            this.SpecificUsers = specificUsers;
            this.AppStatus = appStatus;
            this.Etoken = etoken;
        }

        /// <summary>
        /// Gets or sets the extension type.
        /// </summary>
        public ExtensionType Type { get; set; }

        /// <summary>
        /// Gets or sets the extension scope.
        /// </summary>
        public ExtensionInstallScope Scope { get; set; }

        /// <summary>
        /// Gets or sets the extension manifest stream.
        /// </summary>
        public Stream ManifestStream { get; set; }

        /// <summary>
        /// Gets or sets the asset ID for Office Marketplace.
        /// </summary>
        public string MarketplaceAssetID { get; set; }

        /// <summary>
        /// Gets or sets the content market for Office Marketplace.
        /// </summary>
        public string MarketplaceContentMarket { get; set; }

        /// <summary>
        /// Gets or sets the app status
        /// </summary>
        public string AppStatus { get; set; }

        /// <summary>
        /// Gets or sets the etoken
        /// </summary>
        public string Etoken { get; set; }

        /// <summary>
        /// Gets or sets the value indicating whether extension is available.
        /// </summary>
        public bool IsAvailable { get; set; }

        /// <summary>
        /// Gets or sets the value indicating whether extension is available.
        /// </summary>
        public bool IsMandatory { get; set; }

        /// <summary>
        /// Gets or sets the value indicating whether extension is enabled by default.
        /// </summary>
        public bool IsEnabledByDefault { get; set; }

        /// <summary>
        /// Gets or sets the extension ProvidedTo value.
        /// </summary>
        public ClientExtensionProvidedTo ProvidedTo { get; set; }

        /// <summary>
        /// Gets or sets the user list this extension is provided to.
        /// </summary>
        public StringList SpecificUsers { get; set; }

        /// <summary>
        /// Reads attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            string value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionType);
            if (!string.IsNullOrEmpty(value))
            {
                this.Type = reader.ReadAttributeValue<ExtensionType>(XmlAttributeNames.ClientExtensionType);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionScope);
            if (!string.IsNullOrEmpty(value))
            {
                this.Scope = reader.ReadAttributeValue<ExtensionInstallScope>(XmlAttributeNames.ClientExtensionScope);
            }
            
            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceAssetID);
            if (!string.IsNullOrEmpty(value))
            {
                this.MarketplaceAssetID = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionMarketplaceAssetID);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceContentMarket);
            if (!string.IsNullOrEmpty(value))
            {
                this.MarketplaceContentMarket = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionMarketplaceContentMarket);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionAppStatus);
            if (!string.IsNullOrEmpty(value))
            {
                this.AppStatus = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionAppStatus);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionEtoken);
            if (!string.IsNullOrEmpty(value))
            {
                this.Etoken = reader.ReadAttributeValue<string>(XmlAttributeNames.ClientExtensionEtoken);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionIsAvailable);
            if (!string.IsNullOrEmpty(value))
            {
                this.IsAvailable = reader.ReadAttributeValue<bool>(XmlAttributeNames.ClientExtensionIsAvailable);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionIsMandatory);
            if (!string.IsNullOrEmpty(value))
            {
                this.IsMandatory = reader.ReadAttributeValue<bool>(XmlAttributeNames.ClientExtensionIsMandatory);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionIsEnabledByDefault);
            if (!string.IsNullOrEmpty(value))
            {
                this.IsEnabledByDefault = reader.ReadAttributeValue<bool>(XmlAttributeNames.ClientExtensionIsEnabledByDefault);
            }

            value = reader.ReadAttributeValue(XmlAttributeNames.ClientExtensionProvidedTo);
            if (!string.IsNullOrEmpty(value))
            {
                this.ProvidedTo = reader.ReadAttributeValue<ClientExtensionProvidedTo>(XmlAttributeNames.ClientExtensionProvidedTo);
            }
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionType, this.Type);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionScope, this.Scope);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceAssetID, this.MarketplaceAssetID);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionMarketplaceContentMarket, this.MarketplaceContentMarket);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionAppStatus, this.AppStatus);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionEtoken, this.Etoken);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionIsAvailable, this.IsAvailable);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionIsMandatory, this.IsMandatory);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionIsEnabledByDefault, this.IsEnabledByDefault);
            writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionProvidedTo, this.ProvidedTo);
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.Manifest:
                    this.ManifestStream = new MemoryStream();
                    reader.ReadBase64ElementValue(this.ManifestStream);
                    this.ManifestStream.Position = 0;
                    return true;
                
                case XmlElementNames.ClientExtensionSpecificUsers:
                    this.SpecificUsers = new StringList();
                    this.SpecificUsers.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ClientExtensionSpecificUsers);
                    return true;
                
                default:
                    return base.TryReadElementFromXml(reader);
            }
        }
    
        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (null != this.SpecificUsers)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.ClientExtensionSpecificUsers);
                this.SpecificUsers.WriteElementsToXml(writer);
                writer.WriteEndElement();
            }

            if (null != this.ManifestStream)
            {                
                if (this.ManifestStream.CanSeek)
                {
                    this.ManifestStream.Position = 0;
                }

                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Manifest);
                writer.WriteBase64ElementValue(this.ManifestStream);
                writer.WriteEndElement();
            }
        }
    }
}
