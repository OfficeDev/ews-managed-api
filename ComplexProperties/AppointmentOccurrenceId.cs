// ---------------------------------------------------------------------------
// <copyright file="AppointmentOccurrenceId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
