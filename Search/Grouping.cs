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
// <summary>Defines the Grouping class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents grouping options in item search operations.
    /// </summary>
    public sealed class Grouping : ISelfValidate, IJsonSerializable
    {
        private SortDirection sortDirection = SortDirection.Ascending;
        private PropertyDefinitionBase groupOn;
        private PropertyDefinitionBase aggregateOn;
        private AggregateType aggregateType;

        /// <summary>
        /// Validates this grouping.
        /// </summary>
        private void InternalValidate()
        {
            EwsUtilities.ValidateParam(this.GroupOn, "GroupOn");
            EwsUtilities.ValidateParam(this.AggregateOn, "AggregateOn");
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Grouping"/> class.
        /// </summary>
        public Grouping()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Grouping"/> class.
        /// </summary>
        /// <param name="groupOn">The property to group on.</param>
        /// <param name="sortDirection">The sort direction.</param>
        /// <param name="aggregateOn">The property to aggregate on.</param>
        /// <param name="aggregateType">The type of aggregate to calculate.</param>
        public Grouping(
            PropertyDefinitionBase groupOn,
            SortDirection sortDirection,
            PropertyDefinitionBase aggregateOn,
            AggregateType aggregateType)
            : this()
        {
            EwsUtilities.ValidateParam(groupOn, "groupOn");
            EwsUtilities.ValidateParam(aggregateOn, "aggregateOn");

            this.groupOn = groupOn;
            this.sortDirection = sortDirection;
            this.aggregateOn = aggregateOn;
            this.aggregateType = aggregateType;
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, XmlElementNames.GroupBy);
            writer.WriteAttributeValue(XmlAttributeNames.Order, this.SortDirection);

            this.GroupOn.WriteToXml(writer);

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.AggregateOn);
            writer.WriteAttributeValue(XmlAttributeNames.Aggregate, this.AggregateType);

            this.AggregateOn.WriteToXml(writer);

            writer.WriteEndElement(); // AggregateOn

            writer.WriteEndElement(); // GroupBy
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            JsonObject jsonGrouping = new JsonObject();

            jsonGrouping.Add(XmlAttributeNames.Order, this.SortDirection);

            throw new NotImplementedException();

            ////return jsonGrouping;
        }

        /// <summary>
        /// Gets or sets the sort direction.
        /// </summary>
        public SortDirection SortDirection
        {
            get { return this.sortDirection; }
            set { this.sortDirection = value; }
        }

        /// <summary>
        /// Gets or sets the property to group on.
        /// </summary>
        public PropertyDefinitionBase GroupOn
        {
            get { return this.groupOn; }
            set { this.groupOn = value; }
        }

        /// <summary>
        /// Gets or sets the property to aggregate on.
        /// </summary>
        public PropertyDefinitionBase AggregateOn
        {
            get { return this.aggregateOn; }
            set { this.aggregateOn = value; }
        }

        /// <summary>
        /// Gets or sets the types of aggregate to calculate.
        /// </summary>
        public AggregateType AggregateType
        {
            get { return this.aggregateType; }
            set { this.aggregateType = value; }
        }

        #region ISelfValidate Members

        /// <summary>
        /// Implements ISelfValidate.Validate. Validates this grouping.
        /// </summary>
        void ISelfValidate.Validate()
        {
            this.InternalValidate();
        }

        #endregion
    }
}