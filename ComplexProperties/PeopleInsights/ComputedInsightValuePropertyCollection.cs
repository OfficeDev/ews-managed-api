// ---------------------------------------------------------------------------
// <copyright file="ComputedInsightValuePropertyCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements the class for computed insight value property collection.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    
    /// <summary>
    /// Represents a collection of computed insight values.
    /// </summary>
    public sealed class ComputedInsightValuePropertyCollection : ComplexPropertyCollection<ComputedInsightValueProperty>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ComputedInsightValuePropertyCollection"/> class.
        /// </summary>
        internal ComputedInsightValuePropertyCollection()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComputedInsightValuePropertyCollection"/> class.
        /// </summary>
        /// <param name="collection">The collection of objects to include.</param>
        internal ComputedInsightValuePropertyCollection(IEnumerable<ComputedInsightValueProperty> collection)
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
        /// <returns>ComputedInsightValueProperty.</returns>
        internal override ComputedInsightValueProperty CreateComplexProperty(string xmlElementName)
        {
            return new ComputedInsightValueProperty();
        }

        /// <summary>
        /// Gets the name of the collection item XML element.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        /// <returns>XML element name.</returns>
        internal override string GetCollectionItemXmlElementName(ComputedInsightValueProperty complexProperty)
        {
            return XmlElementNames.Property;
        }
    }
}
