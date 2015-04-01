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
    /// Represents a property definition for properties of type TimeZoneInfo.
    /// </summary>
    internal class StartTimeZonePropertyDefinition : TimeZonePropertyDefinition
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="StartTimeZonePropertyDefinition"/> class.
        /// </summary>
        /// <param name="xmlElementName">Name of the XML element.</param>
        /// <param name="uri">The URI.</param>
        /// <param name="flags">The flags.</param>
        /// <param name="version">The version.</param>
        internal StartTimeZonePropertyDefinition(
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
        /// Registers associated internal properties.
        /// </summary>
        /// <param name="properties">The list in which to add the associated properties.</param>
        internal override void RegisterAssociatedInternalProperties(List<PropertyDefinition> properties)
        {
            base.RegisterAssociatedInternalProperties(properties);

            properties.Add(AppointmentSchema.MeetingTimeZone);
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
            object value = propertyBag[this];

            if (value != null)
            {
                if (writer.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
                {
                    ExchangeService service = writer.Service as ExchangeService;
                    if (service != null && service.Exchange2007CompatibilityMode == false)
                    {
                        MeetingTimeZone meetingTimeZone = new MeetingTimeZone((TimeZoneInfo)value);
                        meetingTimeZone.WriteToXml(writer, XmlElementNames.MeetingTimeZone);
                    }
                }
                else
                {
                    base.WritePropertyValueToXml(
                        writer,
                        propertyBag,
                        isUpdateOperation);
                }
            }
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer)
        {
            if (writer.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
            {
                AppointmentSchema.MeetingTimeZone.WriteToXml(writer);
            }
            else
            {
                base.WriteToXml(writer);
            }
        }

        /// <summary>
        /// Determines whether the specified flag is set.
        /// </summary>
        /// <param name="flag">The flag.</param>
        /// <param name="version">Requested version.</param>
        /// <returns>
        ///     <c>true</c> if the specified flag is set; otherwise, <c>false</c>.
        /// </returns>
        internal override bool HasFlag(PropertyDefinitionFlags flag, ExchangeVersion? version)
        {
            if (version.HasValue && (version.Value == ExchangeVersion.Exchange2007_SP1))
            {
                return AppointmentSchema.MeetingTimeZone.HasFlag(flag, version);
            }
            else
            {
                return base.HasFlag(flag, version);
            }
        }
    }
}