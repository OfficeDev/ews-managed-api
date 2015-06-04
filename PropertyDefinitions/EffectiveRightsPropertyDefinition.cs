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
        /// Gets the property type.
        /// </summary>
        public override Type Type
        {
            get { return typeof(EffectiveRights); }
        }
    }
}