#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the ContainedPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents contained property definition.
    /// </summary>
    /// <typeparam name="TComplexProperty">The type of the complex property.</typeparam>
    internal class ContainedPropertyDefinition<TComplexProperty> : ComplexPropertyDefinition<TComplexProperty> where TComplexProperty : ComplexProperty, new()
    {
        private string containedXmlElementName;

        /// <summary>
        /// Initializes a new instance of the <see cref="ContainedPropertyDefinition&lt;TComplexProperty&gt;"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="containedXmlElementName">Name of the contained XML element.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        /// <param name="propertyCreationDelegate">Delegate used to create instances of ComplexProperty.</param>
        internal ContainedPropertyDefinition(
            string xmlElementName,
            string uri,
            string containedXmlElementName,
            PropertyDefinitionFlags flags,
            ExchangeVersion version,
            CreateComplexPropertyDelegate<TComplexProperty> propertyCreationDelegate)
            : base(xmlElementName, uri, flags, version, propertyCreationDelegate)
        {
            this.containedXmlElementName = containedXmlElementName;
        }

        /// <summary>
        /// Load from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override void InternalLoadFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            reader.ReadStartElement(XmlNamespace.Types, this.containedXmlElementName);

            base.InternalLoadFromXml(reader, propertyBag);

            reader.ReadEndElementIfNecessary(XmlNamespace.Types, this.containedXmlElementName);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
        internal override void WritePropertyValueToXml(
            EwsServiceXmlWriter writer,
            PropertyBag propertyBag,
            bool isUpdateOperation)
        {
            ComplexProperty complexProperty = (ComplexProperty)propertyBag[this];

            if (complexProperty != null)
            {
                writer.WriteStartElement(XmlNamespace.Types, this.XmlElementName);

                complexProperty.WriteToXml(writer, this.containedXmlElementName);

                writer.WriteEndElement(); // this.XmlElementName
            }
        }
    }
}