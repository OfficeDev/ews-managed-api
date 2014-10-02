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
// <summary>Defines the GetServerTimeZonesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the response to a GetServerTimeZones request.
    /// </summary>
    internal class GetServerTimeZonesResponse : ServiceResponse
    {
        private Collection<TimeZoneInfo> timeZones = new Collection<TimeZoneInfo>();

        /// <summary>
        /// Initializes a new instance of the <see cref="GetServerTimeZonesResponse"/> class.
        /// </summary>
        internal GetServerTimeZonesResponse()
            : base()
        {
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.TimeZoneDefinitions);

            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.TimeZoneDefinition))
                    {
                        TimeZoneDefinition timeZoneDefinition = new TimeZoneDefinition();
                        timeZoneDefinition.LoadFromXml(reader);

                        this.timeZones.Add(timeZoneDefinition.ToTimeZoneInfo());
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.TimeZoneDefinitions));
            }
        }

        /// <summary>
        /// Gets the time zones returned by the associated GetServerTimeZones request.
        /// </summary>
        /// <value>The time zones.</value>
        public Collection<TimeZoneInfo> TimeZones
        {
            get { return this.timeZones; }
        }
    }
}
