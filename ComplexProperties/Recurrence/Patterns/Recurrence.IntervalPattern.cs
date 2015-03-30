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
    using System.ComponentModel;
    using System.Text;

    /// <content>
    /// Contains nested type Recurrence.IntervalPattern.
    /// </content>
    public abstract partial class Recurrence
    {
        /// <summary>
        /// Represents a recurrence pattern where each occurrence happens at a specific interval after the previous one.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public abstract class IntervalPattern : Recurrence
        {
            private int interval = 1;

            /// <summary>
            /// Initializes a new instance of the <see cref="IntervalPattern"/> class.
            /// </summary>
            internal IntervalPattern()
            {
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="IntervalPattern"/> class.
            /// </summary>
            /// <param name="startDate">The start date.</param>
            /// <param name="interval">The interval.</param>
            internal IntervalPattern(DateTime startDate, int interval)
                : base(startDate)
            {
                if (interval < 1)
                {
                    throw new ArgumentOutOfRangeException("interval", Strings.IntervalMustBeGreaterOrEqualToOne);
                }

                this.Interval = interval;
            }

            /// <summary>
            /// Write properties to XML.
            /// </summary>
            /// <param name="writer">The writer.</param>
            internal override void InternalWritePropertiesToXml(EwsServiceXmlWriter writer)
            {
                base.InternalWritePropertiesToXml(writer);

                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Interval,
                    this.Interval);
            }

            /// <summary>
            /// Patterns to json.
            /// </summary>
            /// <param name="service">The service.</param>
            /// <returns></returns>
            internal override JsonObject PatternToJson(ExchangeService service)
            {
                JsonObject jsonPattern = new JsonObject();

                jsonPattern.AddTypeParameter(this.XmlElementName);
                jsonPattern.Add(XmlElementNames.Interval, this.Interval);

                return jsonPattern;
            }

            /// <summary>
            /// Tries to read element from XML.
            /// </summary>
            /// <param name="reader">The reader.</param>
            /// <returns>True if appropriate element was read.</returns>
            internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
                if (base.TryReadElementFromXml(reader))
                {
                    return true;
                }
                else
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.Interval:
                            this.interval = reader.ReadElementValue<int>();
                            return true;
                        default:
                            return false;
                    }
                }
            }

            /// <summary>
            /// Loads from json.
            /// </summary>
            /// <param name="jsonProperty">The json property.</param>
            /// <param name="service">The service.</param>
            internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
            {
                base.LoadFromJson(jsonProperty, service);

                foreach (string key in jsonProperty.Keys)
                {
                    switch (key)
                    {
                        case XmlElementNames.Interval:
                            this.interval = jsonProperty.ReadAsInt(key);
                            break;
                        default:
                            break;
                    }
                }
            }

            /// <summary>
            /// Gets or sets the interval between occurrences. 
            /// </summary>
            public int Interval
            {
                get
                {
                    return this.interval;
                }

                set
                {
                    if (value < 1)
                    {
                        throw new ArgumentOutOfRangeException("value", Strings.IntervalMustBeGreaterOrEqualToOne);
                    }

                    this.SetFieldValue<int>(ref this.interval, value);
                }
            }
        }
    }
}