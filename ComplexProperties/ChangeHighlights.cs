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
    /// Encapsulates information on the changehighlights of a meeting request.
    /// </summary>
    public sealed class ChangeHighlights : ComplexProperty
    {
        private bool hasLocationChanged;
        private string location;
        private bool hasStartTimeChanged;
        private DateTime start;
        private bool hasEndTimeChanged;
        private DateTime end;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChangeHighlights"/> class.
        /// </summary>
        internal ChangeHighlights()
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
                case XmlElementNames.HasLocationChanged:
                    this.hasLocationChanged = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Location:
                    this.location = reader.ReadElementValue();
                    return true;
                case XmlElementNames.HasStartTimeChanged:
                    this.hasStartTimeChanged = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Start:
                    this.start = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.HasEndTimeChanged:
                    this.hasEndTimeChanged = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.End:
                    this.end = reader.ReadElementValueAsDateTime().Value;
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
                    case XmlElementNames.HasLocationChanged:
                        this.hasLocationChanged = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.Location:
                        this.location = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.HasStartTimeChanged:
                        this.hasStartTimeChanged = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.Start:
                        this.start = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    case XmlElementNames.HasEndTimeChanged:
                        this.hasEndTimeChanged = jsonProperty.ReadAsBool(key);
                        break;
                    case XmlElementNames.End:
                        this.end = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(key)).Value;
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the location has changed.
        /// </summary>
        public bool HasLocationChanged
        {
            get { return this.hasLocationChanged; }
        }

        /// <summary>
        /// Gets the old location
        /// </summary>
        public string Location
        {
            get { return this.location; }
        }

        /// <summary>
        /// Gets a value indicating whether the the start time has changed.
        /// </summary>
        public bool HasStartTimeChanged
        {
            get { return this.hasStartTimeChanged; }
        }

        /// <summary>
        /// Gets the old start date and time of the meeting.
        /// </summary>
        public DateTime Start
        {
            get { return this.start; }
        }

        /// <summary>
        /// Gets a value indicating whether the the end time has changed.
        /// </summary>
        public bool HasEndTimeChanged
        {
            get { return this.hasEndTimeChanged; }
        }

        /// <summary>
        /// Gets the old end date and time of the meeting.
        /// </summary>
        public DateTime End
        {
            get { return this.end; }
        }
    }
}