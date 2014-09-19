// ---------------------------------------------------------------------------
// <copyright file="PropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents the definition of a folder or item property.
    /// </summary>
    public abstract class PropertyDefinition : ServiceObjectPropertyDefinition
    {
        private string xmlElementName;
        private PropertyDefinitionFlags flags;
        private string name;
        private ExchangeVersion version;

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="version">The version.</param>
        internal PropertyDefinition(
            string xmlElementName,
            string uri,
            ExchangeVersion version)
            : base(uri)
        {
            this.xmlElementName = xmlElementName;
            this.flags = PropertyDefinitionFlags.None;
            this.version = version;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal PropertyDefinition(
            string xmlElementName,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base()
        {
            this.xmlElementName = xmlElementName;
            this.flags = flags;
            this.version = version;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal PropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : this(xmlElementName, uri, version)
        {
            this.flags = flags;
        }

        /// <summary>
        /// Determines whether the specified flag is set.
        /// </summary>
        /// <param name="flag">The flag.</param>
        /// <returns>
        ///     <c>true</c> if the specified flag is set; otherwise, <c>false</c>.
        /// </returns>
        internal bool HasFlag(PropertyDefinitionFlags flag)
        {
            return this.HasFlag(flag, null);
        }

        /// <summary>
        /// Determines whether the specified flag is set.
        /// </summary>
        /// <param name="flag">The flag.</param>
        /// <param name="version">Requested version.</param>
        /// <returns>
        ///     <c>true</c> if the specified flag is set; otherwise, <c>false</c>.
        /// </returns>
        internal virtual bool HasFlag(PropertyDefinitionFlags flag, ExchangeVersion? version)
        {
            return (this.flags & flag) == flag;
        }

        /// <summary>
        /// Registers associated internal properties.
        /// </summary>
        /// <param name="properties">The list in which to add the associated properties.</param>
        internal virtual void RegisterAssociatedInternalProperties(List<PropertyDefinition> properties)
        {
        }

        /// <summary>
        /// Gets a list of associated internal properties.
        /// </summary>
        /// <returns>A list of PropertyDefinition objects.</returns>
        /// <remarks>
        /// This is a hack. It is here (currently) solely to help the API
        /// register the MeetingTimeZone property definition that is internal.
        /// </remarks>
        internal List<PropertyDefinition> GetAssociatedInternalProperties()
        {
            List<PropertyDefinition> properties = new List<PropertyDefinition>();

            this.RegisterAssociatedInternalProperties(properties);

            return properties;
        }

        /// <summary>
        /// Gets the minimum Exchange version that supports this property.
        /// </summary>
        /// <value>The version.</value>
        public override ExchangeVersion Version
        {
            get { return this.version; }
        }

        /// <summary>
        /// Gets a value indicating whether this property definition is for a nullable type (ref, int?, bool?...).
        /// </summary>
        internal virtual bool IsNullable
        {
            get { return true; }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal abstract void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag);

        /// <summary>
        /// Loads the property value from json.
        /// </summary>
        /// <param name="value">The JSON value.  Can be a JsonObject, string, number, bool, array, or null.</param>
        /// <param name="service">The service.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal abstract void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag);

        /// <summary>
        /// Writes the property value to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="isUpdateOperation">Indicates whether the context is an update operation.</param>
        internal abstract void WritePropertyValueToXml(
            EwsServiceXmlWriter writer,
            PropertyBag propertyBag,
            bool isUpdateOperation);

        /// <summary>
        /// Writes the json value.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="service">The service.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        internal abstract void WriteJsonValue(JsonObject jsonObject, PropertyBag propertyBag, ExchangeService service, bool isUpdateOperation);

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <value>The name of the XML element.</value>
        internal string XmlElementName
        {
            get { return this.xmlElementName; }
        }

        /// <summary>
        /// Gets the name of the property.
        /// </summary>
        public string Name
        {
            get
            {
                // Name is initialized at read time for all PropertyDefinition instances using Reflection.
                if (string.IsNullOrEmpty(this.name))
                {
                    ServiceObjectSchema.InitializeSchemaPropertyNames();
                }

                return this.name;
            }

            internal set
            {
                this.name = value;
            }
        }

        /// <summary>
        /// Gets the property definition's printable name.
        /// </summary>
        /// <returns>
        /// The property definition's printable name.
        /// </returns>
        internal override string GetPrintableName()
        {
            return this.Name;
        }
    }
}