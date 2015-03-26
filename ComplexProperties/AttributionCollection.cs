// ---------------------------------------------------------------------------
// <copyright file="AttributionCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements an attribution collection.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents a collection of attributions
    /// </summary>
    public sealed class AttributionCollection : ComplexPropertyCollection<Attribution>
    {
        /// <summary>
        /// XML element name
        /// </summary>
        private readonly string collectionItemXmlElementName;

        /// <summary>
        /// Creates a new instance of the <see cref="AttributionCollection"/> class.
        /// </summary>
        internal AttributionCollection()
            : this(XmlElementNames.Attribution)
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="AttributionCollection"/> class.
        /// </summary>
        /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
        internal AttributionCollection(string collectionItemXmlElementName)
            : base()
        {
            EwsUtilities.ValidateParam(collectionItemXmlElementName, "collectionItemXmlElementName");
            this.collectionItemXmlElementName = collectionItemXmlElementName;
        }

        /// <summary>
        /// Adds an attribution to the collection.
        /// </summary>
        /// <param name="attribution">Attributions to be added</param>
        public void Add(Attribution attribution)
        {
            this.InternalAdd(attribution);
        }

        /// <summary>
        /// Adds multiple attributions to the collection.
        /// </summary>
        /// <param name="attributions">Attributions to be added</param>
        public void AddRange(IEnumerable<Attribution> attributions)
        {
            if (attributions != null)
            {
                foreach (Attribution attribution in attributions)
                {
                    this.Add(attribution);
                }
            }
        }

        /// <summary>
        /// Clears the collection.
        /// </summary>
        public void Clear()
        {
            this.InternalClear();
        }

        /// <summary>
        /// Creates an attribution object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">Attribution XML node name</param>
        /// <returns>The attribution object created</returns>
        internal override Attribution CreateComplexProperty(string xmlElementName)
        {
            EwsUtilities.ValidateParam(xmlElementName, "xmlElementName");
            if (xmlElementName == this.collectionItemXmlElementName)
            {
                return new Attribution();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns></returns>
        internal override Attribution CreateDefaultComplexProperty()
        {
            return new Attribution();
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided attribution object.
        /// </summary>
        /// <param name="attribution">The attribution object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided attribution object.</returns>
        internal override string GetCollectionItemXmlElementName(Attribution attribution)
        {
            return this.collectionItemXmlElementName;
        }

        /// <summary>
        /// Determine whether we should write collection to XML or not.
        /// </summary>
        /// <returns>Always true, even if the collection is empty.</returns>
        internal override bool ShouldWriteToRequest()
        {
            return true;
        }
    }
}
