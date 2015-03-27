// ---------------------------------------------------------------------------
// <copyright file="RequestedUnifiedGroupsSet.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RequestedUnifiedGroupsSet class.</summary>
//-----------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Data.Groups
{
    /// <summary>
    /// Defines the RequestedUnifiedGroupsSet class.
    /// </summary>
    public sealed class RequestedUnifiedGroupsSet : ComplexProperty, ISelfValidate, IJsonSerializable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedUnifiedGroupsSet"/> class.
        /// </summary>
        public RequestedUnifiedGroupsSet()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RequestedUnifiedGroupsSet"/> class.
        /// </summary>
        /// <param name="filterType">The filterType for the list of groups to be returned</param>
        /// <param name="sortType">The sort type for the list of groups to be returned</param>
        /// <param name="sortDirection">The sort direction for list of groups to be returned</param>
        public RequestedUnifiedGroupsSet(
            UnifiedGroupsFilterType filterType,
            UnifiedGroupsSortType sortType,
            SortDirection sortDirection)
        {
            this.FilterType = filterType;
            this.SortType = sortType;
            this.SortDirection = sortDirection;
        }

        /// <summary>
        /// Gets or sets the sort type for the list of groups to be returned
        /// </summary>
        public UnifiedGroupsSortType SortType { get; set; }

        /// <summary>
        /// Gets or sets the filter Type for the list of groups to be returned
        /// </summary>
        public UnifiedGroupsFilterType FilterType { get; set; }

        /// <summary>
        /// Gets or sets the Sort Direction for the list of groups to be returned.
        /// </summary>
        public SortDirection SortDirection { get; set; }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(EwsServiceXmlWriter writer, string xmlElementName)
        {
            writer.WriteStartElement(XmlNamespace.Types, xmlElementName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SortType, this.SortType.ToString());
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FilterType, this.FilterType.ToString());
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.SortDirection, this.SortDirection.ToString());

            writer.WriteEndElement();
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();
            jsonProperty.Add(XmlElementNames.SortType, this.SortType.ToString());
            jsonProperty.Add(XmlElementNames.FilterType, this.FilterType.ToString());
            jsonProperty.Add(XmlElementNames.SortDirection, this.SortDirection.ToString());

            return jsonProperty;
        }
    }
}
