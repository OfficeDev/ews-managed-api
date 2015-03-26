// ---------------------------------------------------------------------------
// <copyright file="AttributedStringCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements an attributed string collection.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents a collection of attributed strings
    /// </summary>
    public sealed class AttributedStringCollection : ComplexPropertyCollection<AttributedString>
    {
        /// <summary>
        /// Collection parent XML element name
        /// </summary>
        private readonly string collectionItemXmlElementName;

        /// <summary>
        /// Creates a new instance of the <see cref="AttributedStringCollection"/> class.
        /// </summary>
        internal AttributedStringCollection()
            : this(XmlElementNames.StringAttributedValue)
        {
        }

        /// <summary>
        /// Creates a new instance of the <see cref="AttributedStringCollection"/> class.
        /// </summary>
        /// <param name="collectionItemXmlElementName">Name of the collection item XML element.</param>
        internal AttributedStringCollection(string collectionItemXmlElementName)
            : base()
        {
            EwsUtilities.ValidateParam(collectionItemXmlElementName, "collectionItemXmlElementName"); 
            this.collectionItemXmlElementName = collectionItemXmlElementName;
        }

        /// <summary>
        /// Adds an attributed string to the collection.
        /// </summary>
        /// <param name="attributedString">Attributed string to be added</param>
        public void Add(AttributedString attributedString)
        {
            this.InternalAdd(attributedString);
        }

        /// <summary>
        /// Adds multiple attributed strings to the collection.
        /// </summary>
        /// <param name="attributedStrings">Attributed strings to be added</param>
        public void AddRange(IEnumerable<AttributedString> attributedStrings)
        {
            if (attributedStrings != null)
            {
                foreach (AttributedString attributedString in attributedStrings)
                {
                    this.Add(attributedString);
                }
            }
        }

        /// <summary>
        /// Adds an attributed string to the collection.
        /// </summary>
        /// <param name="stringValue">The SMTP address used to initialize the e-mail address.</param>
        /// <returns>An AttributedString object initialized with the provided SMTP address.</returns>
        public AttributedString Add(string stringValue)
        {
            AttributedString attributedString = new AttributedString(stringValue);

            this.Add(attributedString);

            return attributedString;
        }

        /// <summary>
        /// Adds a string value and list of attributions
        /// </summary>
        /// <param name="stringValue">String value of the attributed string being added</param>
        /// <param name="attributions">Attributions of the attributed string being added</param>
        /// <returns>The added attributedString object</returns>
        public AttributedString Add(string stringValue, IList<string> attributions)
        {
            AttributedString attributedString = new AttributedString(stringValue, attributions);

            this.Add(attributedString);

            return attributedString;
        }

        /// <summary>
        /// Clears the collection.
        /// </summary>
        public void Clear()
        {
            this.InternalClear();
        }

        /// <summary>
        /// Removes an attributed string from the collection.
        /// </summary>
        /// <param name="attributedString">Attributed string to be removed</param>
        /// <returns>Whether succeeded</returns>
        public bool Remove(AttributedString attributedString)
        {
            EwsUtilities.ValidateParam(attributedString, "attributedString");

            return this.InternalRemove(attributedString);
        }

        /// <summary>
        /// Creates an AttributedString object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to create the attributed string object</param>
        /// <returns>An AttributedString object</returns>
        internal override AttributedString CreateComplexProperty(string xmlElementName)
        {
            EwsUtilities.ValidateParam(xmlElementName, "xmlElementName");
            if (xmlElementName == this.collectionItemXmlElementName)
            {
                return new AttributedString();
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
        internal override AttributedString CreateDefaultComplexProperty()
        {
            return new AttributedString();
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided AttributedString object.
        /// </summary>
        /// <param name="attributedString">The AttributedString object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided AttributedString object.</returns>
        internal override string GetCollectionItemXmlElementName(AttributedString attributedString)
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
