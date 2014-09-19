// ---------------------------------------------------------------------------
// <copyright file="EffectiveRightsPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EffectiveRightsPropertyDefinition class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents effective rights property definition.
    /// </summary>
    internal sealed class EffectiveRightsPropertyDefinition : PropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="EffectiveRightsPropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal EffectiveRightsPropertyDefinition(
            string xmlElementName,
            string uri,
            PropertyDefinitionFlags flags,
            ExchangeVersion version)
            : base(
                xmlElementName,
                uri,
                flags,
                version)
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="propertyBag">The property bag.</param>
        internal override sealed void LoadPropertyValueFromXml(EwsServiceXmlReader reader, PropertyBag propertyBag)
        {
            EffectiveRights value = EffectiveRights.None;

            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, this.XmlElementName);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement())
                    {
                        switch (reader.LocalName)
                        {
                            case XmlElementNames.CreateAssociated:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.CreateAssociated;
                                }
                                break;
                            case XmlElementNames.CreateContents:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.CreateContents;
                                }
                                break;
                            case XmlElementNames.CreateHierarchy:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.CreateHierarchy;
                                }
                                break;
                            case XmlElementNames.Delete:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.Delete;
                                }
                                break;
                            case XmlElementNames.Modify:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.Modify;
                                }
                                break;
                            case XmlElementNames.Read:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.Read;
                                }
                                break;
                            case XmlElementNames.ViewPrivateItems:
                                if (reader.ReadElementValue<bool>())
                                {
                                    value |= EffectiveRights.ViewPrivateItems;
                                }
                                break;
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Types, this.XmlElementName));
            }

            propertyBag[this] = value;
        }

        internal override void LoadPropertyValueFromJson(object value, ExchangeService service, PropertyBag propertyBag)
        {
            EffectiveRights effectiveRightsValue = EffectiveRights.None;
            JsonObject jsonObject = value as JsonObject;

            if (jsonObject != null)
            {
                foreach (string key in jsonObject.Keys)
                {
                    switch (key)
                    {
                        case XmlElementNames.CreateAssociated:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.CreateAssociated;
                            }
                            break;
                        case XmlElementNames.CreateContents:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.CreateContents;
                            }
                            break;
                        case XmlElementNames.CreateHierarchy:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.CreateHierarchy;
                            }
                            break;
                        case XmlElementNames.Delete:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.Delete;
                            }
                            break;
                        case XmlElementNames.Modify:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.Modify;
                            }
                            break;
                        case XmlElementNames.Read:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.Read;
                            }
                            break;
                        case XmlElementNames.ViewPrivateItems:
                            if (jsonObject.ReadAsBool(key))
                            {
                                effectiveRightsValue |= EffectiveRights.ViewPrivateItems;
                            }
                            break;
                    }
                }
            }

            propertyBag[this] = value;
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
            // EffectiveRights is a read-only property, no need to implement this.
        }

        /// <summary>
        /// Writes the json value.
        /// </summary>
        /// <param name="jsonObject">The json object.</param>
        /// <param name="propertyBag">The property bag.</param>
        /// <param name="service">The service.</param>
        /// <param name="isUpdateOperation">if set to <c>true</c> [is update operation].</param>
        internal override void WriteJsonValue(JsonObject jsonObject, PropertyBag propertyBag, ExchangeService service, bool isUpdateOperation)
        {
            // EffectiveRights is a read-only property, no need to implement this.
        }

        /// <summary>
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(EffectiveRights); }
        }
    }
}