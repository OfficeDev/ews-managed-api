// ---------------------------------------------------------------------------
// <copyright file="DeletedOccurrenceInfoCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeletedOccurrenceInfoCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of deleted occurrence objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class DeletedOccurrenceInfoCollection : ComplexPropertyCollection<DeletedOccurrenceInfo>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OccurrenceInfoCollection"/> class.
        /// </summary>
        internal DeletedOccurrenceInfoCollection()
        {
        }

        /// <summary>
        /// Creates the complex property.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <returns>OccurenceInfo instance.</returns>
        internal override DeletedOccurrenceInfo CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == XmlElementNames.DeletedOccurrence)
            {
                return new DeletedOccurrenceInfo();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns></returns>
        internal override DeletedOccurrenceInfo CreateDefaultComplexProperty()
        {
            return new DeletedOccurrenceInfo();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(DeletedOccurrenceInfo complexProperty)
        {
            return XmlElementNames.Occurrence;
        }
    }
}
