// ---------------------------------------------------------------------------
// <copyright file="ClientAppMetadata.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ClientAppMetadata class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.Xml;

    /// <summary>
    /// Represents a ClientAppMetadata object.
    /// </summary>
    public sealed class ClientAppMetadata : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ClientAppMetadata"/> class.
        /// </summary>
        internal ClientAppMetadata()
            : base()
        {
            this.Namespace = XmlNamespace.Types;
        }

        /// <summary>
        /// The End node url for the app.
        /// </summary>
        public string EndNodeUrl
        {
            get;
            private set;
        }

        /// <summary>
        /// The action url for the app.
        /// </summary>
        public string ActionUrl
        {
            get;
            private set;
        }

        /// <summary>
        /// The app status for the app.
        /// </summary>
        public string AppStatus
        {
            get;
            private set;
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
                case XmlElementNames.EndNodeUrl:
                    this.EndNodeUrl = reader.ReadElementValue<string>();
                    return true;
                case XmlElementNames.ActionUrl:
                    this.ActionUrl = reader.ReadElementValue<string>();
                    return true;
                case XmlElementNames.AppStatus:
                    this.AppStatus = reader.ReadElementValue<string>();
                    return true;
                default:
                    return false;
            }
        }
    }
}
