// ---------------------------------------------------------------------------
// <copyright file="SearchFolderParameters.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SearchFolderParameters class.</summary>
//-----------------------------------------------------------------------

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
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.BaseFolderIds:
                        this.RootFolderIds.InternalClear();
                        ((IJsonCollectionDeserializer)this.RootFolderIds).CreateFromJsonCollection(jsonProperty.ReadAsArray(key), service);
                        break;
                    case XmlElementNames.Restriction:
                        JsonObject restriction = jsonProperty.ReadAsJsonObject(key);
                        this.searchFilter = SearchFilter.LoadSearchFilterFromJson(restriction.ReadAsJsonObject(XmlElementNames.Item), service);
                        break;
                    case XmlAttributeNames.Traversal:
                        this.Traversal = jsonProperty.ReadEnumValue<SearchFolderTraversal>(key);
                        break;
                    default:
                        break;
                }
            }
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
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlAttributeNames.Traversal, this.Traversal);
            jsonProperty.Add(XmlElementNames.BaseFolderIds, this.RootFolderIds.InternalToJson(service));

            if (this.SearchFilter != null)
            {
                JsonObject restriction = new JsonObject();
                restriction.Add(XmlElementNames.Item, this.SearchFilter.InternalToJson(service));
                jsonProperty.Add(XmlElementNames.Restriction, restriction);
            }

            return jsonProperty;
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
