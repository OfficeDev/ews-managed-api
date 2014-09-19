// ---------------------------------------------------------------------------
// <copyright file="TaskSuggestionCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TaskSuggestionCollection class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a collection of TaskSuggestion objects.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class TaskSuggestionCollection : ComplexPropertyCollection<TaskSuggestion>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TaskSuggestionCollection"/> class.
        /// </summary>
        internal TaskSuggestionCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TaskSuggestionCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal TaskSuggestionCollection(IEnumerable<TaskSuggestion> collection)
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
        /// <returns>TaskSuggestion.</returns>
        internal override TaskSuggestion CreateComplexProperty(string xmlElementName)
        {
            return new TaskSuggestion();
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>TaskSuggestion.</returns>
        internal override TaskSuggestion CreateDefaultComplexProperty()
        {
            return new TaskSuggestion();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(TaskSuggestion complexProperty)
        {
            return XmlElementNames.NlgTaskSuggestion;
        }
    }
}
