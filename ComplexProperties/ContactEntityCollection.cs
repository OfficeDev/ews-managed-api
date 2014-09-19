// ---------------------------------------------------------------------------
// <copyright file="ContactEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of ContactEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ContactEntityCollection : ComplexPropertyCollection<ContactEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ContactEntityCollection"/> class.
        /// </summary>
        internal ContactEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal ContactEntityCollection(IEnumerable<ContactEntity> collection)
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
        /// <returns>ContactEntity.</returns>
        internal override ContactEntity CreateComplexProperty(string xmlElementName)
        {
            return new ContactEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>ContactEntity.</returns>
        internal override ContactEntity CreateDefaultComplexProperty()
        {
            return new ContactEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(ContactEntity complexProperty)
        {
            return XmlElementNames.NlgContact;
        }
    }
}
