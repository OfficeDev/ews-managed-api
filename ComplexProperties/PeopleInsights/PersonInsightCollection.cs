/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents the collection of person insights. 
    /// </summary>
    public sealed class PersonInsightCollection : ComplexPropertyCollection<PersonInsight>
    {
        /// <summary>
        /// XML element name
        /// </summary>
        private readonly string collectionItemXmlElementName;

        /// <summary>
        /// Creates a new instance of the <see cref="PersonInsightCollection"/> class.
        /// </summary>
        internal PersonInsightCollection()
            : this(XmlElementNames.Insight)
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="PersonInsightCollection"/> class.
        /// </summary>
        /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
        internal PersonInsightCollection(string collectionItemXmlElementName)
            : base()
        {
            this.collectionItemXmlElementName = collectionItemXmlElementName;
        }

        /// <summary>
        /// Adds an insight to the collection.
        /// </summary>
        /// <param name="insight">The insight to add.</param>
        public void Add(PersonInsight insight)
        {
            this.InternalAdd(insight);
        }

        /// <summary>
        /// Adds multiple person insights to the collection.
        /// </summary>
        /// <param name="insights">The collection of person insights to add.</param>
        public void AddRange(IEnumerable<PersonInsight> insights)
        {
            if (insights != null)
            {
                foreach (PersonInsight insight in insights)
                {
                    this.Add(insight);
                }
            }
        }

        /// <summary>
        /// Clears the collection.
        /// </summary>
        public void Clear()
        {
            this.InternalClear();
        }

        /// <summary>
        /// Removes a person insight from the collection.
        /// </summary>
        /// <param name="index">The index of the person insight to remove.</param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= this.Count)
            {
                throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
            }

            this.InternalRemoveAt(index);
        }

        /// <summary>
        /// Removes a person insight from the collection.
        /// </summary>
        /// <param name="insight">The person sight to remove.</param>
        /// <returns>Whether removed from the collection</returns>
        public bool Remove(PersonInsight insight)
        {
            EwsUtilities.ValidateParam(insight, "insight");

            return this.InternalRemove(insight);
        }

        /// <summary>
        /// Creates a PersonInsight object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to create the PersonInsight.</param>
        /// <returns>A PersonInsight object.</returns>
        internal override PersonInsight CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == this.collectionItemXmlElementName)
            {
                return new PersonInsight();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided PersonInsight object.
        /// </summary>
        /// <param name="insight">The PersonInsight object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided PersonInsight object.</returns>
        internal override string GetCollectionItemXmlElementName(PersonInsight insight)
        {
            return this.collectionItemXmlElementName;
        }

        /// <summary>
        /// Determine whether we should write collection to XML or not.
        /// </summary>
        /// <returns>Always true, even if the collection is empty.</returns>
        internal override bool ShouldWriteToRequest()
        {
            return true;
        }
    }
}