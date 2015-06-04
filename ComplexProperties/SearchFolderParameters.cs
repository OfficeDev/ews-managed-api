/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the parameters associated with a search folder.
    /// </summary>
    public sealed class SearchFolderParameters : ComplexProperty
    {
        private SearchFolderTraversal traversal;
        private FolderIdCollection rootFolderIds = new FolderIdCollection();
        private SearchFilter searchFilter;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchFolderParameters"/> class.
        /// </summary>
        internal SearchFolderParameters()
            : base()
        {
            this.rootFolderIds.OnChange += this.PropertyChanged;
        }

        /// <summary>
        /// Property changed.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        private void PropertyChanged(ComplexProperty complexProperty)
        {
            this.Changed();
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
                case XmlElementNames.BaseFolderIds:
                    this.RootFolderIds.InternalClear();
                    this.RootFolderIds.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Restriction:
                    reader.Read();
                    this.searchFilter = SearchFilter.LoadFromXml(reader);
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
        {
            this.Traversal = reader.ReadAttributeValue<SearchFolderTraversal>(XmlAttributeNames.Traversal);
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Traversal, this.Traversal);
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.SearchFilter != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Restriction);
                this.SearchFilter.WriteToXml(writer);
                writer.WriteEndElement(); // Restriction
            }

            this.RootFolderIds.WriteToXml(writer, XmlElementNames.BaseFolderIds);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal void Validate()
        {
            // Search folder must have at least one root folder id.
            if (this.RootFolderIds.Count == 0)
            {
                throw new ServiceValidationException(Strings.SearchParametersRootFolderIdsEmpty);
            }

            // Validate the search filter
            if (this.SearchFilter != null)
            {
                this.SearchFilter.InternalValidate();
            }
        }

        /// <summary>
        /// Gets or sets the traversal mode for the search folder.
        /// </summary>
        public SearchFolderTraversal Traversal
        {
            get { return this.traversal; }
            set { this.SetFieldValue<SearchFolderTraversal>(ref this.traversal, value); }
        }

        /// <summary>
        /// Gets the list of root folders the search folder searches in.
        /// </summary>
        public FolderIdCollection RootFolderIds
        {
            get { return this.rootFolderIds; }
        }

        /// <summary>
        /// Gets or sets the search filter associated with the search folder. Available search filter classes include
        /// SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.
        /// </summary>
        public SearchFilter SearchFilter
        {
            get
            {
                return this.searchFilter;
            }

            set
            {
                if (this.searchFilter != null)
                {
                    this.searchFilter.OnChange -= this.PropertyChanged;
                }

                this.SetFieldValue<SearchFilter>(ref this.searchFilter, value);

                if (this.searchFilter != null)
                {
                    this.searchFilter.OnChange += this.PropertyChanged;
                }
            }
        }
    }
}