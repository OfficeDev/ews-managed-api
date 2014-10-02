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
// <summary>Defines the AppointmentOccurrenceId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents the Id of an occurrence of a recurring appointment.
    /// </summary>
    public sealed class AppointmentOccurrenceId : ItemId
    {
        /// <summary>
        /// Index of the occurrence.
        /// </summary>
        private int occurrenceIndex;

        /// <summary>
        /// Initializes a new instance of the <see cref="AppointmentOccurrenceId"/> class.
        /// </summary>
        /// <param name="recurringMasterUniqueId">The Id of the recurring master the Id represents an occurrence of.</param>
        /// <param name="occurrenceIndex">The index of the occurrence.</param>
        public AppointmentOccurrenceId(string recurringMasterUniqueId, int occurrenceIndex)
            : base(recurringMasterUniqueId)
        {
            this.OccurrenceIndex = occurrenceIndex;
        }

        /// <summary>
        /// Gets or sets the index of the occurrence. Note that the occurrence index starts at one not zero.
        /// </summary>
        public int OccurrenceIndex
        {
            get
            { 
                return this.occurrenceIndex;
            }

            set
            {
                // The occurence index has to be positive integer.
                if (value < 1)
                {
                    throw new ArgumentException(Strings.OccurrenceIndexMustBeGreaterThanZero);
                }

                this.occurrenceIndex = value;
            }
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.OccurrenceItemId;
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.RecurringMasterId, this.UniqueId);
            writer.WriteAttributeValue(XmlAttributeNames.InstanceIndex, this.OccurrenceIndex);
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.AddTypeParameter(this.GetXmlElementName());
            jsonProperty.Add(XmlAttributeNames.RecurringMasterId, this.UniqueId);
            jsonProperty.Add(XmlAttributeNames.InstanceIndex, this.OccurrenceIndex);

            return jsonProperty;
        }
    }
}
