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
    /// Encapsulates information on the occurrence of a recurring appointment.
    /// </summary>
    public sealed class Flag : ComplexProperty
    {
        private ItemFlagStatus flagStatus;
        private DateTime startDate;
        private DateTime dueDate;
        private DateTime completeDate;

        /// <summary>
        /// Initializes a new instance of the <see cref="Flag"/> class.
        /// </summary>
        public Flag()
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.FlagStatus:
                    this.flagStatus = reader.ReadElementValue<ItemFlagStatus>();
                    return true;
                case XmlElementNames.StartDate:
                    this.startDate = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.DueDate:
                    this.dueDate = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.CompleteDate:
                    this.completeDate = reader.ReadElementValueAsDateTime().Value;
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.FlagStatus:
                        this.flagStatus = jsonProperty.ReadEnumValue<ItemFlagStatus>(key);
                        break;
                    case XmlElementNames.StartDate:
                        this.startDate = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    case XmlElementNames.DueDate:
                        this.dueDate = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    case XmlElementNames.CompleteDate:
                        this.completeDate = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FlagStatus, this.FlagStatus);

            if (this.FlagStatus == ItemFlagStatus.Flagged && this.StartDate != null && this.DueDate != null)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.StartDate, this.StartDate);
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DueDate, this.DueDate);
            }
            else if (this.FlagStatus == ItemFlagStatus.Complete && this.CompleteDate != null)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.CompleteDate, this.CompleteDate);
            }
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
            JsonObject jsonObject = new JsonObject();

            jsonObject.Add(XmlElementNames.FlagStatus, this.FlagStatus);
            if (this.FlagStatus == ItemFlagStatus.Flagged && this.StartDate != null && this.DueDate != null)
            {
                jsonObject.Add(XmlElementNames.StartDate, this.StartDate);
                jsonObject.Add(XmlElementNames.DueDate, this.DueDate);
            }
            else if (this.FlagStatus == ItemFlagStatus.Complete && this.CompleteDate != null)
            {
                jsonObject.Add(XmlElementNames.CompleteDate, this.CompleteDate);
            }

            return jsonObject;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal void Validate()
        {
            EwsUtilities.ValidateParam(this.flagStatus, "FlagStatus");
        }

        /// <summary>
        /// Gets or sets the flag status.
        /// </summary>
        public ItemFlagStatus FlagStatus
        {
            get
            {
                return this.flagStatus;
            }

            set
            {
                this.SetFieldValue<ItemFlagStatus>(ref this.flagStatus, value);
            }
        }

        /// <summary>
        /// Gets the start date.
        /// </summary>
        public DateTime StartDate
        {
            get 
            { 
                return this.startDate; 
            }

            set
            {
                this.SetFieldValue<DateTime>(ref this.startDate, value);
            }
        }

        /// <summary>
        /// Gets the due date.
        /// </summary>
        public DateTime DueDate
        {
            get 
            { 
                return this.dueDate; 
            }

            set
            {
                this.SetFieldValue<DateTime>(ref this.dueDate, value);
            }
        }

        /// <summary>
        /// Gets the complete date.
        /// </summary>
        public DateTime CompleteDate
        {
            get 
            { 
                return this.completeDate; 
            }

            set
            {
                this.SetFieldValue<DateTime>(ref this.completeDate, value);
            }
        }
    }
}