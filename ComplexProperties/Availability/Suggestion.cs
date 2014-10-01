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
// <summary>Defines the Suggestion class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.Text;

    /// <summary>
    /// Represents a suggestion for a specific date.
    /// </summary>
    public sealed class Suggestion : ComplexProperty
    {
        private DateTime date;
        private SuggestionQuality quality;
        private Collection<TimeSuggestion> timeSuggestions = new Collection<TimeSuggestion>();

        /// <summary>
        /// Initializes a new instance of the <see cref="Suggestion"/> class.
        /// </summary>
        internal Suggestion()
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
                case XmlElementNames.Date:
                    // The date that is returned by Availability is unscoped. 
                    DateTime tempDate = DateTime.Parse(reader.ReadElementValue(), CultureInfo.InvariantCulture);

                    if (tempDate.Kind != DateTimeKind.Unspecified)
                    {
                        this.date = new DateTime(tempDate.Ticks, DateTimeKind.Unspecified);
                    }
                    else
                    {
                        this.date = tempDate;
                    }

                    return true;
                case XmlElementNames.DayQuality:
                    this.quality = reader.ReadElementValue<SuggestionQuality>();
                    return true;
                case XmlElementNames.SuggestionArray:
                    if (!reader.IsEmptyElement)
                    {
                        do
                        {
                            reader.Read();

                            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Suggestion))
                            {
                                TimeSuggestion timeSuggestion = new TimeSuggestion();

                                timeSuggestion.LoadFromXml(reader, reader.LocalName);

                                this.timeSuggestions.Add(timeSuggestion);
                            }
                        }
                        while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.SuggestionArray));
                    }

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
            this.date = DateTime.Parse(jsonProperty.ReadAsString(XmlElementNames.Date));
            this.quality = jsonProperty.ReadEnumValue<SuggestionQuality>(XmlElementNames.DayQuality);

            foreach (object jsonSuggestion in jsonProperty.ReadAsArray(XmlElementNames.SuggestionArray))
            {
                TimeSuggestion timeSuggestion = new TimeSuggestion();

                timeSuggestion.LoadFromJson(jsonSuggestion as JsonObject, service);

                this.timeSuggestions.Add(timeSuggestion);
            }
        }
        
        /// <summary>
        /// Gets the date and time of the suggestion.
        /// </summary>
        public DateTime Date
        {
            get { return this.date; }
        }

        /// <summary>
        /// Gets the quality of the suggestion.
        /// </summary>
        public SuggestionQuality Quality
        {
            get { return this.quality; }
        }

        /// <summary>
        /// Gets a collection of suggested times within the suggested day.
        /// </summary>
        public Collection<TimeSuggestion> TimeSuggestions
        {
            get { return this.timeSuggestions; }
        }
    }
}