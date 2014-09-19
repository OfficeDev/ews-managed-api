// ---------------------------------------------------------------------------
// <copyright file="GroupMemberPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GroupMemberPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the definition of the GroupMember property.
    /// </summary>
    internal sealed class GroupMemberPropertyDefinition : ServiceObjectPropertyDefinition
    {
        /// <summary>
        /// FieldUri of IndexedFieldURI for a group member.
        /// </summary>
        private const string FieldUri = "distributionlist:Members:Member";

        /// <summary>
        /// Member key.
        /// Maps to the Index attribute of IndexedFieldURI element.
        /// </summary>
        private string key;

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupMemberPropertyDefinition"/> class.
        /// </summary>
        /// <param name="key">The member's key.</param>
        public GroupMemberPropertyDefinition(string key)
            : base(FieldUri)
        {
            this.key = key;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="GroupMemberPropertyDefinition"/> class without key.
        /// </summary>
        internal GroupMemberPropertyDefinition()
            : base(FieldUri)
        {
        }

        /// <summary>
        /// Gets or sets the member's key.
        /// </summary>
        public string Key
        {
            get
            {
                return this.key;
            }

            set
            {
                this.key = value;
            }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.IndexedFieldURI;
        }

        protected override string GetJsonType()
        {
            return JsonNames.PathToIndexedFieldType;
        }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            base.WriteAttributesToXml(writer);
            writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, this.Key);
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonPropertyDefinition">The json property definition.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonPropertyDefinition, ExchangeService service)
        {
            base.AddJsonProperties(jsonPropertyDefinition, service);
            jsonPropertyDefinition.Add(XmlAttributeNames.FieldIndex, this.Key);
        }

        /// <summary>
        /// Gets the property definition's printable name.
        /// </summary>
        /// <returns>
        /// The property definition's printable name.
        /// </returns>
        internal override string GetPrintableName()
        {
            return string.Format("{0}:{1}", FieldUri, this.Key);
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(string); }
        }
    }
}
