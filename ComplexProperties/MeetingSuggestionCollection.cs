// ---------------------------------------------------------------------------
// <copyright file="MeetingSuggestionCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the MeetingSuggestionCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of MeetingSuggestion objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class MeetingSuggestionCollection : ComplexPropertyCollection<MeetingSuggestion>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingSuggestionCollection"/> class.
        /// </summary>
        internal MeetingSuggestionCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingSuggestionCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal MeetingSuggestionCollection(IEnumerable<MeetingSuggestion> collection)
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
        /// <returns>MeetingSuggestion.</returns>
        internal override MeetingSuggestion CreateComplexProperty(string xmlElementName)
        {
            return new MeetingSuggestion();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>MeetingSuggestion.</returns>
        internal override MeetingSuggestion CreateDefaultComplexProperty()
        {
            return new MeetingSuggestion();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(MeetingSuggestion complexProperty)
        {
            return XmlElementNames.NlgMeetingSuggestion;
        }
    }
}
