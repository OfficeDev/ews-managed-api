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
    using System.Xml;

    /// <summary>
    /// Represents an attributed string, a string with a value and a list of attributions.
    /// </summary>
    public sealed class AttributedString : ComplexProperty
    {
        /// <summary>
        /// Internal attribution store
        /// </summary>
        private List<string> attributionList;

        /// <summary>
        /// String value
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Attribution values
        /// </summary>
        public IList<string> Attributions { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public AttributedString()
            : base()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public AttributedString(string value)
            : this()
        {
            EwsUtilities.ValidateParam(value, "value");
            this.Value = value;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="value">String value</param>
        /// <param name="attributions">A list of attributions</param>
        public AttributedString(string value, IList<string> attributions)
            : this(value)
        {
            if (attributions == null)
            {
                throw new ArgumentNullException("attributions");
            }

            foreach (string s in attributions)
            {
                EwsUtilities.ValidateParam(s, "attributions");
            }

            this.Attributions = attributions;
        }

        /// <summary>
        /// Defines an implicit conversion from a regular string to an attributedString.
        /// </summary>
        /// <param name="value">String value of the attributed string being created</param>
        /// <returns>An attributed string initialized with the specified value</returns>
        public static implicit operator AttributedString(string value)
        {
            return new AttributedString(value);
        }

        /// <summary>
        /// Tries to read an attributed string blob represented in XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether reading succeeded</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.Value:
                    this.Value = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Attributions:
                    return this.LoadAttributionsFromXml(reader);
                default:
                    return false;
            }
        }

        /// <summary>
        /// Read attribution blobs from XML
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether reading succeeded</returns>
        internal bool LoadAttributionsFromXml(EwsServiceXmlReader reader)
        {
            if (!reader.IsEmptyElement)
            {
                string localName = reader.LocalName;
                this.attributionList = new List<string>();

                do
                {
                    reader.Read();
                    if (reader.NodeType == XmlNodeType.Element &&
                        reader.LocalName == XmlElementNames.Attribution)
                    {
                        string s = reader.ReadElementValue();
                        if (!string.IsNullOrEmpty(s))
                        {
                            this.attributionList.Add(s);
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, localName));
                this.Attributions = this.attributionList.ToArray();
            }

            return true;
        }
    }
}