// ---------------------------------------------------------------------------
// <copyright file="NameResolutionCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the NameResolutionCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a list of suggested name resolutions.
    /// </summary>
    public sealed class NameResolutionCollection : IEnumerable<NameResolution>
    {
        private ExchangeService service;
        private bool includesAllResolutions;
        private List<NameResolution> items = new List<NameResolution>();

        /// <summary>
        /// Initializes a new instance of the <see cref="NameResolutionCollection"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal NameResolutionCollection(ExchangeService service)
        {
            EwsUtilities.Assert(
                service != null,
                "NameResolutionSet.ctor",
                "service is null.");

            this.service = service;
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResolutionSet);

            int totalItemsInView = reader.ReadAttributeValue<int>(XmlAttributeNames.TotalItemsInView);
            this.includesAllResolutions = reader.ReadAttributeValue<bool>(XmlAttributeNames.IncludesLastItemInRange);

            for (int i = 0; i < totalItemsInView; i++)
            {
                NameResolution nameResolution = new NameResolution(this);

                nameResolution.LoadFromXml(reader);

                this.items.Add(nameResolution);
            }

            reader.ReadEndElement(XmlNamespace.Messages, XmlElementNames.ResolutionSet);
        }

        /// <summary>
        /// Loads from json array.
        /// </summary>
        /// <param name="jsonProperty">The p.</param>
        /// <param name="service">The service.</param>
        internal void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            int totalItemsInView;
            object[] resolutions;

            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlAttributeNames.TotalItemsInView:
                        totalItemsInView = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlAttributeNames.IncludesLastItemInRange:
                        this.includesAllResolutions = jsonProperty.ReadAsBool(key);
                        break;
                   
                    // This label only exists for Json objects.  The XML doesn't have a "Resolutions"
                    // element.  
                    // This was necessary becaue of the lack of attributes in JSON
                    //
                    case "Resolutions":
                        resolutions = jsonProperty.ReadAsArray(key);
                        foreach (object resolution in resolutions)
                        {
                            JsonObject resolutionProperty = resolution as JsonObject;
                            if (resolutionProperty != null)
                            {
                                NameResolution nameResolution = new NameResolution(this);
                                nameResolution.LoadFromJson(resolutionProperty, service);
                                this.items.Add(nameResolution);
                            }
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets the session.
        /// </summary>
        /// <value>The session.</value>
        internal ExchangeService Session
        {
            get { return this.service; }
        }

        /// <summary>
        /// Gets the total number of elements in the list.
        /// </summary>
        public int Count
        {
            get { return this.items.Count; }
        }

        /// <summary>
        /// Gets a value indicating whether more suggested resolutions are available. ResolveName only returns
        /// a maximum of 100 name resolutions. When IncludesAllResolutions is false, there were more than 100
        /// matching names on the server. To narrow the search, provide a more precise name to ResolveName.
        /// </summary>
        public bool IncludesAllResolutions
        {
            get { return this.includesAllResolutions; }
        }

        /// <summary>
        /// Gets the name resolution at the specified index.
        /// </summary>
        /// <param name="index">The index of the name resolution to get.</param>
        /// <returns>The name resolution at the speicfied index.</returns>
        public NameResolution this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                }

                return this.items[index];
            }
        }

        #region IEnumerable<NameResolution> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<NameResolution> GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.items.GetEnumerator();
        }

        #endregion
    }
}
