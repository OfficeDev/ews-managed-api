// ---------------------------------------------------------------------------
// <copyright file="ContainedPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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