// ---------------------------------------------------------------------------
// <copyright file="EmailAddressEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailAddressEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of EmailAddressEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class EmailAddressEntityCollection : ComplexPropertyCollection<EmailAddressEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressEntityCollection"/> class.
        /// </summary>
        internal EmailAddressEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailAddressEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal EmailAddressEntityCollection(IEnumerable<EmailAddressEntity> collection)
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
        /// <returns>EmailAddressEntity.</returns>
        internal override EmailAddressEntity CreateComplexProperty(string xmlElementName)
        {
            return new EmailAddressEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>EmailAddressEntity.</returns>
        internal override EmailAddressEntity CreateDefaultComplexProperty()
        {
            return new EmailAddressEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(EmailAddressEntity complexProperty)
        {
            return XmlElementNames.NlgEmailAddress;
        }
    }
}
