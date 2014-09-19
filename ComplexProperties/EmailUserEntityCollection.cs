// ---------------------------------------------------------------------------
// <copyright file="EmailUserEntityCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EmailUserEntityCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of EmailUserEntity objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class EmailUserEntityCollection : ComplexPropertyCollection<EmailUserEntity>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EmailUserEntityCollection"/> class.
        /// </summary>
        internal EmailUserEntityCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailUserEntityCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal EmailUserEntityCollection(IEnumerable<EmailUserEntity> collection)
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
        /// <returns>EmailUserEntity.</returns>
        internal override EmailUserEntity CreateComplexProperty(string xmlElementName)
        {
            return new EmailUserEntity();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>EmailUserEntity.</returns>
        internal override EmailUserEntity CreateDefaultComplexProperty()
        {
            return new EmailUserEntity();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(EmailUserEntity complexProperty)
        {
            return XmlElementNames.NlgEmailUser;
        }
    }
}
