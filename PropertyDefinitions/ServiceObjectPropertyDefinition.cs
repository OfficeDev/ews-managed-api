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
    using System.Text;

    /// <summary>
    /// Represents a property definition for a service object.
    /// </summary>
    public abstract class ServiceObjectPropertyDefinition : PropertyDefinitionBase
    {
        private string uri;

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.FieldURI;
        }

        /// <summary>
        /// Gets the type for json.
        /// </summary>
        /// <returns></returns>
        protected override string GetJsonType()
        {
            return JsonNames.PathToUnindexedFieldType;
        }

        /// <summary>
        /// Gets the minimum Exchange version that supports this property.
        /// </summary>
        /// <value>The version.</value>
        public override ExchangeVersion Version
        {
            get { return ExchangeVersion.Exchange2007_SP1; }
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.FieldURI, this.Uri);
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonPropertyDefinition">The json property definition.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonPropertyDefinition, ExchangeService service)
        {
            jsonPropertyDefinition.Add(XmlAttributeNames.FieldURI, this.Uri);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceObjectPropertyDefinition"/> class.
        /// </summary>
        internal ServiceObjectPropertyDefinition()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceObjectPropertyDefinition"/> class.
        /// </summary>
        /// <param name="uri">The URI.</param>
        internal ServiceObjectPropertyDefinition(string uri)
            : base()
        {
            EwsUtilities.Assert(
                !string.IsNullOrEmpty(uri),
                "ServiceObjectPropertyDefinition.ctor",
                "uri is null or empty");

            this.uri = uri;
        }

        /// <summary>
        /// Gets the URI of the property definition.
        /// </summary>
        internal string Uri
        {
            get { return this.uri; }
        }
    }
}