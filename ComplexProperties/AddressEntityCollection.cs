// ---------------------------------------------------------------------------
// <copyright file="AddressEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AddressEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of AddressEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class AddressEntityCollection : ComplexPropertyCollection<AddressEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AddressEntityCollection"/> class.
        /// </summary>
        internal AddressEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="AddressEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal AddressEntityCollection(IEnumerable<AddressEntity> collection)
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
        /// <returns>AddressEntity.</returns>
        internal override AddressEntity CreateComplexProperty(string xmlElementName)
        {
            return new AddressEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>AddressEntity.</returns>
        internal override AddressEntity CreateDefaultComplexProperty()
        {
            return new AddressEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(AddressEntity complexProperty)
        {
            return XmlElementNames.NlgAddress;
        }
    }
}
