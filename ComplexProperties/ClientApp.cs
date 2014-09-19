// ---------------------------------------------------------------------------
// <copyright file="ClientApp.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ClientApp class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents a app in GetAppManifests response.
    /// </summary>
    public sealed class ClientApp : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ClientApp"/> class.
        /// </summary>
        internal ClientApp()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }

        /// <summary>
        /// The manifest for the app.
        /// </summary>
        public XmlDocument Manifest 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Metadata related to the app.
        /// </summary>
        public ClientAppMetadata Metadata
        {
            get;
            internal set;
        }

        /// <summary>
        /// Helper to convert to xml dcouemnt from the current value.
        /// </summary>
        /// <param name="reader">the reader.</param>
        /// <returns>The xml document</returns>
        internal static SafeXmlDocument ReadToXmlDocument(EwsServiceXmlReader reader)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                reader.ReadBase64ElementValue(stream);
                stream.Position = 0;

                SafeXmlDocument manifest = new SafeXmlDocument();
                manifest.Load(stream);
                return manifest;
            }
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
                    this.Manifest = ClientApp.ReadToXmlDocument(reader);
                    return true;

                case XmlElementNames.Metadata:
                    this.Metadata = new ClientAppMetadata();
                    this.Metadata.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Metadata);
                    return true;

                default:
                    return false;
            }
        }
    }
}
