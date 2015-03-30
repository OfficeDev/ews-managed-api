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
    using System.Text;

    /// <summary>
    /// Represents a date range view of appointments in calendar folder search operations.
    /// </summary>
    public sealed class CalendarView : ViewBase
    {
        private ItemTraversal traversal;
        private int? maxItemsReturned;
        private DateTime startDate;
        private DateTime endDate;

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.Traversal, this.Traversal);
        }

        /// <summary>
        /// Writes the search settings to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="groupBy">The group by clause.</param>
        internal override void InternalWriteSearchSettingsToXml(EwsServiceXmlWriter writer, Grouping groupBy)
        {
            // No search settings for calendar views.
        }

        /// <summary>
        /// Writes the grouping to json.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="groupBy"></param>
        /// <returns></returns>
        internal override object WriteGroupingToJson(ExchangeService service, Grouping groupBy)
        {
            // No search settings for calendar views.
            return null;
        }

        /// <summary>
        /// Writes OrderBy property to XML.
        /// </summary>
        /// <param name="writer">The writer</param>
        internal override void WriteOrderByToXml(EwsServiceXmlWriter writer)
        {
            // No OrderBy for calendar views.
        }

        /// <summary>
        /// Adds the json properties.
        /// </summary>
        /// <param name="jsonRequest">The json request.</param>
        /// <param name="service">The service.</param>
        internal override void AddJsonProperties(JsonObject jsonRequest, ExchangeService service)
        {
            jsonRequest.Add(XmlAttributeNames.Traversal, this.Traversal);
        }

        /// <summary>
        /// Gets the type of service object this view applies to.
        /// </summary>
        /// <returns>A ServiceObjectType value.</returns>
        internal override ServiceObjectType GetServiceObjectType()
        {
            return ServiceObjectType.Item;
        }

        /// <summary>
        /// Initializes a new instance of CalendarView.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="endDate">The end date.</param>
        public CalendarView(
            DateTime startDate,
            DateTime endDate)
            : base()
        {
            this.startDate = startDate;
            this.endDate = endDate;
        }

        /// <summary>
        /// Initializes a new instance of CalendarView.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="endDate">The end date.</param>
        /// <param name="maxItemsReturned">The maximum number of items the search operation should return.</param>
        public CalendarView(
            DateTime startDate,
            DateTime endDate,
            int maxItemsReturned)
            : this(startDate, endDate)
        {
            this.MaxItemsReturned = maxItemsReturned;
        }

        /// <summary>
        /// Validate instance.
        /// </summary>
        /// <param name="request">The request using this view.</param>
        internal override void InternalValidate(ServiceRequestBase request)
        {
            base.InternalValidate(request);

            if (this.endDate < this.StartDate)
            {
                throw new ServiceValidationException(Strings.EndDateMustBeGreaterThanStartDate);
            }
        }

        /// <summary>
        /// Write to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void InternalWriteViewToXml(EwsServiceXmlWriter writer)
        {
            base.InternalWriteViewToXml(writer);

            writer.WriteAttributeValue(XmlAttributeNames.StartDate, this.StartDate);
            writer.WriteAttributeValue(XmlAttributeNames.EndDate, this.EndDate);
        }

        /// <summary>
        /// Internals the write paging to json.
        /// </summary>
        /// <param name="jsonView">The json view.</param>
        /// <param name="service">The service.</param>
        internal override void InternalWritePagingToJson(JsonObject jsonView, ExchangeService service)
        {
            base.InternalWritePagingToJson(jsonView, service);

            jsonView.Add(XmlAttributeNames.StartDate, this.StartDate);
            jsonView.Add(XmlAttributeNames.EndDate, this.EndDate);
        }

        /// <summary>
        /// Gets the name of the view XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetViewXmlElementName()
        {
            return XmlElementNames.CalendarView;
        }

        /// <summary>
        /// Gets the name of the view json type.
        /// </summary>
        /// <returns></returns>
        internal override string GetViewJsonTypeName()
        {
            return "CalendarPageView";
        }

        /// <summary>
        /// Gets the maximum number of items or folders the search operation should return.
        /// </summary>
        /// <returns>The maximum number of items the search operation should return.
        /// </returns>
        internal override int? GetMaxEntriesReturned()
        {
            return this.MaxItemsReturned;
        }

        /// <summary>
        /// Gets or sets the start date.
        /// </summary>
        public DateTime StartDate
        {
            get { return this.startDate; }
            set { this.startDate = value; }
        }

        /// <summary>
        /// Gets or sets the end date.
        /// </summary>
        public DateTime EndDate
        {
            get { return this.endDate; }
            set { this.endDate = value; }
        }

        /// <summary>
        /// The maximum number of items the search operation should return.
        /// </summary>
        public int? MaxItemsReturned
        {
            get
            {
                return this.maxItemsReturned;
            }

            set
            {
                if (value.HasValue)
                {
                    if (value.Value <= 0)
                    {
                        throw new ArgumentException(Strings.ValueMustBeGreaterThanZero);
                    }
                }

                this.maxItemsReturned = value;
            }
        }

        /// <summary>
        /// Gets or sets the search traversal mode. Defaults to ItemTraversal.Shallow.
        /// </summary>
        public ItemTraversal Traversal
        {
            get { return this.traversal; }
            set { this.traversal = value; }
        }
    }
}