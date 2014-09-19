// ---------------------------------------------------------------------------
// <copyright file="ContactPhoneEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactPhoneEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of ContactPhoneEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ContactPhoneEntityCollection : ComplexPropertyCollection<ContactPhoneEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ContactPhoneEntityCollection"/> class.
        /// </summary>
        internal ContactPhoneEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactPhoneEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal ContactPhoneEntityCollection(IEnumerable<ContactPhoneEntity> collection)
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
        /// <returns>ContactPhoneEntity.</returns>
        internal override ContactPhoneEntity CreateComplexProperty(string xmlElementName)
        {
            return new ContactPhoneEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>ContactPhoneEntity.</returns>
        internal override ContactPhoneEntity CreateDefaultComplexProperty()
        {
            return new ContactPhoneEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(ContactPhoneEntity complexProperty)
        {
            return XmlElementNames.NlgPhone;
        }
    }
}
