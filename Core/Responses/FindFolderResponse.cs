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
// <summary>Defines the FindFolderResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents the response to a folder search operation.
    /// </summary>
    internal sealed class FindFolderResponse : ServiceResponse
    {
        private FindFoldersResults results = new FindFoldersResults();
        private PropertySet propertySet;

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.RootFolder);

            this.results.TotalCount = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
            this.results.MoreAvailable = !reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

            // Ignore IndexedPagingOffset attribute if MoreAvailable is false.
            this.results.NextPageOffset = results.MoreAvailable ? reader.ReadNullableAttributeValue<int>(XmlAttributeNames.IndexedPagingOffset) : null;

            reader.ReadStartElement(XmlNamespace.Types, XmlElementNames.Folders);
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        Folder folder = EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(reader.Service, reader.LocalName);

                        if (folder == null)
                        {
                            reader.SkipCurrentElement();
                        }
                        else
                        {
                            folder.LoadFromXml(
                                        reader,
                                        true, /* clearPropertyBag */
                                        this.propertySet,
                                        true  /* summaryPropertiesOnly */);

                            this.results.Folders.Add(folder);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.Folders));
            }

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.RootFolder);
        }

        /// <summary>
        /// Reads response elements from Json.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="service">The service.</param>
        internal override void ReadElementsFromJson(JsonObject responseObject, ExchangeService service)
        {
            JsonObject rootFolder = responseObject.ReadAsJsonObject(XmlElementNames.RootFolder);
            this.results.TotalCount = rootFolder.ReadAsInt(XmlAttributeNames.TotalItemsInView);
            this.results.MoreAvailable = rootFolder.ReadAsBool(XmlAttributeNames.IncludesLastItemInRange);

            // Ignore IndexedPagingOffset attribute if MoreAvailable is false.
            if (results.MoreAvailable)
            {
                if (rootFolder.ContainsKey(XmlAttributeNames.IndexedPagingOffset))
                {
                    this.results.NextPageOffset = rootFolder.ReadAsInt(XmlAttributeNames.IndexedPagingOffset);
                }
                else
                {
                    this.results.NextPageOffset = null;
                }
            }

            if (rootFolder.ContainsKey(XmlElementNames.Folders))
            {
                List<Folder> folders = new EwsServiceJsonReader(service).ReadServiceObjectsCollectionFromJson<Folder>(
                    rootFolder,
                    XmlElementNames.Folders,
                    this.CreateFolderInstance,
                    true,               /* clearPropertyBag */
                    this.propertySet,   /* requestedPropertySet */
                    true);              /* summaryPropertiesOnly */

                folders.ForEach((folder) => this.results.Folders.Add(folder));
            }
        }

        /// <summary>
        /// Creates a folder instance.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>Folder</returns>
        private Folder CreateFolderInstance(ExchangeService service, string xmlElementName)
        {
            return EwsUtilities.CreateEwsObjectFromXmlElementName<Folder>(service, xmlElementName);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FindFolderResponse"/> class.
        /// </summary>
        /// <param name="propertySet">The property set from, the request.</param>
        internal FindFolderResponse(PropertySet propertySet)
            : base()
        {
            this.propertySet = propertySet;

            EwsUtilities.Assert(
                this.propertySet != null,
                "FindFolderResponse.ctor",
                "PropertySet should not be null");
        }

        /// <summary>
        /// Gets the results of the search operation.
        /// </summary>
        public FindFoldersResults Results
        {
            get { return this.results; }
        }
    }
}
