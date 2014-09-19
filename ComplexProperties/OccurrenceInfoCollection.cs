// ---------------------------------------------------------------------------
// <copyright file="OccurrenceInfoCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OccurrenceInfoCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of OccurrenceInfo objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class OccurrenceInfoCollection : ComplexPropertyCollection<OccurrenceInfo>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OccurrenceInfoCollection"/> class.
        /// </summary>
        internal OccurrenceInfoCollection()
        {
        }

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>OccurenceInfo instance.</returns>
        internal override OccurrenceInfo CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == XmlElementNames.Occurrence)
            {
                return new OccurrenceInfo();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>OccurenceInfo instance.</returns>
        internal override OccurrenceInfo CreateDefaultComplexProperty()
        {
            return new OccurrenceInfo();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(OccurrenceInfo complexProperty)
        {
            return XmlElementNames.Occurrence;
        }
    }
}
