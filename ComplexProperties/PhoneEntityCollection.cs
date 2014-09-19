// ---------------------------------------------------------------------------
// <copyright file="PhoneEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PhoneEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of PhoneEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class PhoneEntityCollection : ComplexPropertyCollection<PhoneEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PhoneEntityCollection"/> class.
        /// </summary>
        internal PhoneEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PhoneEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal PhoneEntityCollection(IEnumerable<PhoneEntity> collection)
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
        /// <returns>PhoneEntity.</returns>
        internal override PhoneEntity CreateComplexProperty(string xmlElementName)
        {
            return new PhoneEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>PhoneEntity.</returns>
        internal override PhoneEntity CreateDefaultComplexProperty()
        {
            return new PhoneEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(PhoneEntity complexProperty)
        {
            return XmlElementNames.NlgPhone;
        }
    }
}
