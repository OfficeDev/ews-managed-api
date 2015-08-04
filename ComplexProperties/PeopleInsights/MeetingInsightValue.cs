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
    using System.Collections.Generic;
    using System.Xml;

    /// <summary>
    /// Represents the MeetingInsightValue.
    /// </summary>
    public sealed class MeetingInsightValue : InsightValue
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MeetingInsightValue"/> class.
        /// </summary>
        public MeetingInsightValue()
            : base()
        {
            this.Attendees = new ProfileInsightValueCollection();
        }

        /// <summary>
        /// Gets the Id
        /// </summary>
        public string Id
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the Subject
        /// </summary>
        public string Subject
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the StartUtcTicks
        /// </summary>
        public long StartUtcTicks
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the EndUtcTicks
        /// </summary>
        public long EndUtcTicks
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the Location
        /// </summary>
        public string Location
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the Organizer
        /// </summary>
        public ProfileInsightValue Organizer
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the Attendees
        /// </summary>
        public ProfileInsightValueCollection Attendees
        {
            get;
            internal set;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.InsightSource:
                    this.InsightSource = reader.ReadElementValue<string>();
                    break;
                case XmlElementNames.UpdatedUtcTicks:
                    this.UpdatedUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.Id:
                    this.Id = reader.ReadElementValue();
                    break;
                case XmlElementNames.Subject:
                    this.Subject = reader.ReadElementValue();
                    break;
                case XmlElementNames.StartUtcTicks:
                    this.StartUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.EndUtcTicks:
                    this.EndUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.Location:
                    this.Location = reader.ReadElementValue();
                    break;
                case XmlElementNames.Organizer:
                    this.Organizer = new ProfileInsightValue();
                    this.Organizer.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.Attendees:
                    this.Attendees = new ProfileInsightValueCollection(XmlElementNames.Item);
                    this.Attendees.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Attendees);
                    break;
                default:
                    return false;
            }

            return true;
        }
    }
}