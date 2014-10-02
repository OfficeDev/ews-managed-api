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
