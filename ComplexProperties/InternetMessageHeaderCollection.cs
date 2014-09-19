// ---------------------------------------------------------------------------
// <copyright file="InternetMessageHeaderCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the InternetMessageHeaderCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents a collection of Internet message headers.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class InternetMessageHeaderCollection : ComplexPropertyCollection<InternetMessageHeader>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InternetMessageHeaderCollection"/> class.
        /// </summary>
        internal InternetMessageHeaderCollection()
            : base()
        {
        }

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>InternetMessageHeader instance.</returns>
        internal override InternetMessageHeader CreateComplexProperty(string xmlElementName)
        {
            return new InternetMessageHeader();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>InternetMessageHeader instance.</returns>
        internal override InternetMessageHeader CreateDefaultComplexProperty()
        {
            return new InternetMessageHeader();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(InternetMessageHeader complexProperty)
        {
            return XmlElementNames.InternetMessageHeader;
        }

        /// <summary>
        /// Find a specific header in the collection.
        /// </summary>
        /// <param name="name">The name of the header to locate.</param>
        /// <returns>An InternetMessageHeader representing the header with the specified name; null if no header with the specified name was found.</returns>
        public InternetMessageHeader Find(string name)
        {
            foreach (InternetMessageHeader internetMessageHeader in this)
            {
                if (string.Compare(name, internetMessageHeader.Name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return internetMessageHeader;
                }
            }

            return null;
        }
    }
}
