// ---------------------------------------------------------------------------
// <copyright file="DocumentSharingLocationCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DocumentSharingLocationCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a user setting that is a collection of alternate mailboxes.
    /// </summary>
    public sealed class DocumentSharingLocationCollection
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentSharingLocationCollection"/> class.
        /// </summary>
        internal DocumentSharingLocationCollection()
        {
            this.Entries = new List<DocumentSharingLocation>();
        }

        /// <summary>
        /// Loads instance of DocumentSharingLocationCollection from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>DocumentSharingLocationCollection</returns>
        internal static DocumentSharingLocationCollection LoadFromXml(EwsXmlReader reader)
        {
            DocumentSharingLocationCollection instance = new DocumentSharingLocationCollection();

            do
            {
                reader.Read();

                if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.DocumentSharingLocation))
                {
                    DocumentSharingLocation location = DocumentSharingLocation.LoadFromXml(reader);
                    instance.Entries.Add(location);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DocumentSharingLocations));

            return instance;
        }

        /// <summary>
        /// Gets the collection of alternate mailboxes.
        /// </summary>
        public List<DocumentSharingLocation> Entries
        {
            get; private set;
        }
    }
}
