// ---------------------------------------------------------------------------
// <copyright file="ItemIdCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ItemIdCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of item Ids.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ItemIdCollection : ComplexPropertyCollection<ItemId>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ItemIdCollection"/> class.
        /// </summary>
        internal ItemIdCollection()
            : base()
        {
        }

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>ItemId.</returns>
        internal override ItemId CreateComplexProperty(string xmlElementName)
        {
            return new ItemId();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>ItemId.</returns>
        internal override ItemId CreateDefaultComplexProperty()
        {
            return new ItemId();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(ItemId complexProperty)
        {
            return complexProperty.GetXmlElementName();
        }
    }
}
