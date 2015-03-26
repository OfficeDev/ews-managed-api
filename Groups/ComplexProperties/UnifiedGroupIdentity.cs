// ---------------------------------------------------------------------------
// <copyright file="UnifiedGroupIdentity.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UnifiedGroupIdentity class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Defines the UnifiedGroupIdentity class.
    /// </summary>
    internal sealed class UnifiedGroupIdentity : ComplexProperty, ISelfValidate, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UnifiedGroupIdentity"/>  class
        /// </summary>
        /// <param name="identityType">The identity type</param>
        /// <param name="value">The value assocaited with the identity type</param>
        public UnifiedGroupIdentity(UnifiedGroupIdentityType identityType, string value)
        {
            this.IdentityType = identityType;
            this.Value = value;
        }

        /// <summary>
        /// Gets or sets the IdentityType of the UnifiedGroup
        /// </summary>
        public UnifiedGroupIdentityType IdentityType { get; set; }

        /// <summary>
        /// Gets or sets the value associated with the IdentityType for the UnifiedGroup
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.GroupIdentityType, this.IdentityType.ToString());
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.GroupIdentityValue, this.Value);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();
            jsonProperty.Add(XmlElementNames.GroupIdentityType, this.IdentityType.ToString());
            jsonProperty.Add(XmlElementNames.GroupIdentityValue, this.Value);

            return jsonProperty;
        }
    }
}
