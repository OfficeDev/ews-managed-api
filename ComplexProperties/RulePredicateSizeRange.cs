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
// <summary>Defines the RulePredicateSizeRange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the minimum and maximum size of a message.
    /// </summary>
    public sealed class RulePredicateSizeRange : ComplexProperty
    {
        /// <summary>
        /// Minimum Size.
        /// </summary>
        private int? minimumSize;

        /// <summary>
        /// Mamixmum Size.
        /// </summary>
        private int? maximumSize;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicateSizeRange"/> class.
        /// </summary>
        internal RulePredicateSizeRange()
            : base()
        {
        }

        /// <summary>
        /// Gets or sets the minimum size, in kilobytes. If MinimumSize is set to 
        /// null, no minimum size applies.
        /// </summary>
        public int? MinimumSize
        {
            get
            {
                return this.minimumSize;
            }

            set
            {
                this.SetFieldValue<int?>(ref this.minimumSize, value);
            }
        }

        /// <summary>
        /// Gets or sets the maximum size, in kilobytes. If MaximumSize is set to 
        /// null, no maximum size applies.
        /// </summary>
        public int? MaximumSize
        {
            get
            {
                return this.maximumSize;
            }

            set
            {
                this.SetFieldValue<int?>(ref this.maximumSize, value);
            }
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
                case XmlElementNames.MinimumSize:
                    this.minimumSize = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.MaximumSize:
                    this.maximumSize = reader.ReadElementValue<int>();
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
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.MinimumSize:
                        this.minimumSize = jsonProperty.ReadAsInt(key);
                        break;
                    case XmlElementNames.MaximumSize:
                        this.maximumSize = jsonProperty.ReadAsInt(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (this.MinimumSize.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MinimumSize, this.MinimumSize.Value);
            }
            if (this.MaximumSize.HasValue)
            {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MaximumSize, this.MaximumSize.Value);
            }
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

            if (this.MinimumSize.HasValue)
            {
                jsonProperty.Add(XmlElementNames.MinimumSize, this.MinimumSize.Value);
            }
            if (this.MaximumSize.HasValue)
            {
                jsonProperty.Add(XmlElementNames.MaximumSize, this.MaximumSize.Value);
            }

            return jsonProperty;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();
            if (this.minimumSize.HasValue &&
                this.maximumSize.HasValue &&
                this.minimumSize.Value > this.maximumSize.Value)
            {
                throw new ServiceValidationException("MinimumSize cannot be larger than MaximumSize.");
            }
        }
    }
}
