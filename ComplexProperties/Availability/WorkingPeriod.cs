#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the WorkingPeriod class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Text;

    /// <summary>
    /// Represents a working period.
    /// </summary>
    internal sealed class WorkingPeriod : ComplexProperty
    {
        private Collection<DayOfTheWeek> daysOfWeek = new Collection<DayOfTheWeek>();
        private TimeSpan startTime;
        private TimeSpan endTime;

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkingPeriod"/> class.
        /// </summary>
        internal WorkingPeriod()
            : base()
        {
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if appropriate element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.DayOfWeek:
                    EwsUtilities.ParseEnumValueList<DayOfTheWeek>(
                        this.daysOfWeek,
                        reader.ReadElementValue(),
                        ' ');
                    return true;
                case XmlElementNames.StartTimeInMinutes:
                    this.startTime = TimeSpan.FromMinutes(reader.ReadElementValue<int>());
                    return true;
                case XmlElementNames.EndTimeInMinutes:
                    this.endTime = TimeSpan.FromMinutes(reader.ReadElementValue<int>());
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
                    case XmlElementNames.DayOfWeek:
                        EwsUtilities.ParseEnumValueList<DayOfTheWeek>(
                            this.daysOfWeek,
                            jsonProperty.ReadAsString(key),
                            ' ');
                        break;
                    case XmlElementNames.StartTimeInMinutes:
                        this.startTime = TimeSpan.FromMinutes(jsonProperty.ReadAsInt(key));
                        break;
                    case XmlElementNames.EndTimeInMinutes:
                        this.endTime = TimeSpan.FromMinutes(jsonProperty.ReadAsInt(key));
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Gets a collection of work days.
        /// </summary>
        internal Collection<DayOfTheWeek> DaysOfWeek
        {
            get { return this.daysOfWeek; }
        }

        /// <summary>
        /// Gets the start time of the period.
        /// </summary>
        internal TimeSpan StartTime
        {
            get { return this.startTime; }
        }

        /// <summary>
        /// Gets the end time of the period.
        /// </summary>
        internal TimeSpan EndTime
        {
            get { return this.endTime; }
        }
    }
}
