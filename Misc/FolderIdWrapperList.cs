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
// <summary>Defines the FolderIdWrapperList class and dependant classes.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a list a abstracted folder Ids.
    /// </summary>
    internal class FolderIdWrapperList : IEnumerable<AbstractFolderIdWrapper>
    {
        /// <summary>
        /// List of <see cref="Microsoft.Exchange.WebServices.Data.AbstractFolderIdWrapper"/>.
        /// </summary>
        private List<AbstractFolderIdWrapper> ids = new List<AbstractFolderIdWrapper>();

        /// <summary>
        /// Adds the specified folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        internal void Add(Folder folder)
        {
            this.ids.Add(new FolderWrapper(folder));
        }

        /// <summary>
        /// Adds the range.
        /// </summary>
        /// <param name="folders">The folders.</param>
        internal void AddRange(IEnumerable<Folder> folders)
        {
            if (folders != null)
            {
                foreach (Folder folder in folders)
                {
                    this.Add(folder);
                }
            }
        }

        /// <summary>
        /// Adds the specified folder id.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        internal void Add(FolderId folderId)
        {
            this.ids.Add(new FolderIdWrapper(folderId));
        }

        /// <summary>
        /// Adds the range of folder ids.
        /// </summary>
        /// <param name="folderIds">The folder ids.</param>
        internal void AddRange(IEnumerable<FolderId> folderIds)
        {
            if (folderIds != null)
            {
                foreach (FolderId folderId in folderIds)
                {
                    this.Add(folderId);
                }
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsNamesapce">The ews namesapce.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal void WriteToXml(
            EwsServiceXmlWriter writer,
            XmlNamespace ewsNamesapce,
            string xmlElementName)
        {
            if (this.Count > 0)
            {
                writer.WriteStartElement(ewsNamesapce, xmlElementName);

                foreach (AbstractFolderIdWrapper folderIdWrapper in this.ids)
                {
                    folderIdWrapper.WriteToXml(writer);
                }

                writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal object InternalToJson(ExchangeService service)
        {
            List<object> jsonArray = new List<object>();

            foreach (AbstractFolderIdWrapper folderIdWrapper in this.ids)
            {
                jsonArray.Add(((IJsonSerializable)folderIdWrapper).ToJson(service));
            }

            return jsonArray.ToArray();
        }

        /// <summary>
        /// Gets the id count.
        /// </summary>
        /// <value>The count.</value>
        internal int Count
        {
            get { return this.ids.Count; }
        }

        /// <summary>
        /// Gets the <see cref="Microsoft.Exchange.WebServices.Data.AbstractFolderIdWrapper"/> at the specified index.
        /// </summary>
        /// <param name="index">the index</param>
        internal AbstractFolderIdWrapper this[int index]
        {
            get { return this.ids[index]; }
        }

        /// <summary>
        /// Validates list of folderIds against a specified request version.
        /// </summary>
        /// <param name="version">The version.</param>
        internal void Validate(ExchangeVersion version)
        {
            foreach (AbstractFolderIdWrapper folderIdWrapper in this.ids)
            {
                folderIdWrapper.Validate(version);
            }
        }

        #region IEnumerable<AbstractFolderIdWrapper> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<AbstractFolderIdWrapper> GetEnumerator()
        {
            return this.ids.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.ids.GetEnumerator();
        }

        #endregion
    }
}
