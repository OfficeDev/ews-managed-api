// ---------------------------------------------------------------------------
// <copyright file="ServiceObjectPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceObjectPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

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
