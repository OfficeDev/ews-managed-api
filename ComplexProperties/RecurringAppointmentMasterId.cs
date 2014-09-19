// ---------------------------------------------------------------------------
// <copyright file="RecurringAppointmentMasterId.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RecurringAppointmentMasterId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the Id of an occurrence of a recurring appointment.
    /// </summary>
    public sealed class RecurringAppointmentMasterId : ItemId
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RecurringAppointmentMasterId"/> class.
        /// </summary>
        /// <param name="occurrenceId">The Id of an occurrence in the recurring series.</param>
        public RecurringAppointmentMasterId(string occurrenceId)
            : base(occurrenceId)
        {
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.RecurringMasterItemId;
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.OccurrenceId, this.UniqueId);
            writer.WriteAttributeValue(XmlAttributeNames.ChangeKey, this.ChangeKey);
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
            jsonProperty.Add(XmlAttributeNames.OccurrenceId, this.UniqueId);
            jsonProperty.Add(XmlAttributeNames.ChangeKey, this.ChangeKey);

            return jsonProperty;
        }
    }
}
