// ---------------------------------------------------------------------------
// <copyright file="UrlEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UrlEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of UrlEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class UrlEntityCollection : ComplexPropertyCollection<UrlEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UrlEntityCollection"/> class.
        /// </summary>
        internal UrlEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UrlEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal UrlEntityCollection(IEnumerable<UrlEntity> collection)
        {
            if (collection != null)
            {
                collection.ForEach(this.InternalAdd);
            }
        }

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>UrlEntity.</returns>
        internal override UrlEntity CreateComplexProperty(string xmlElementName)
        {
            return new UrlEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>UrlEntity.</returns>
        internal override UrlEntity CreateDefaultComplexProperty()
        {
            return new UrlEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(UrlEntity complexProperty)
        {
            return XmlElementNames.NlgUrl;
        }
    }
}
