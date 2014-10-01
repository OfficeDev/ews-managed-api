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
// <summary>Defines the DeletedOccurrenceInfo class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Encapsulates information on the deleted occurrence of a recurring appointment.
    /// </summary>
    public class DeletedOccurrenceInfo : ComplexProperty
    {
        /// <summary>
        /// The original start date and time of the deleted occurrence.
        /// </summary>
        /// <remarks>
        /// The EWS schema contains a Start property for deleted occurrences but it's
        /// really the original start date and time of the occurrence.
        /// </remarks>
        private DateTime originalStart;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeletedOccurrenceInfo"/> class.
        /// </summary>
        internal DeletedOccurrenceInfo()
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
                case XmlElementNames.Start:
                    this.originalStart = reader.ReadElementValueAsDateTime().Value;
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            if (jsonProperty.ContainsKey(XmlElementNames.Start))
            {
                this.originalStart = service.ConvertUniversalDateTimeStringToLocalDateTime(jsonProperty.ReadAsString(XmlElementNames.Start)).Value;
            }
        }

        /// <summary>
        /// Gets the original start date and time of the deleted occurrence.
        /// </summary>
        public DateTime OriginalStart
        {
            get { return this.originalStart; }
        }
    }
}
