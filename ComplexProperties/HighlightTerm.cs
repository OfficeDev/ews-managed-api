// ---------------------------------------------------------------------------
// <copyright file="HighlightTerm.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the HighlightTerm class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an AQS highlight term. 
    /// </summary>
    public sealed class HighlightTerm : ComplexProperty
    {
        /// <summary>
        /// Term scope.
        /// </summary>
        private string scope;

        /// <summary>
        /// Term value.
        /// </summary>
        private string value;

        /// <summary>
        /// Initializes a new instance of the <see cref="HighlightTerm"/> class.
        /// </summary>
        internal HighlightTerm()
            : base()
        {
        }

        /// <summary>
        /// Gets term scope.
        /// </summary>
        public string Scope
        {
            get
            {
                return this.scope;
            }
        }

        /// <summary>
        /// Gets term value.
        /// </summary>
        public string Value
        {
            get
            {
                return this.value;
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
                case XmlElementNames.HighlightTermScope:
                    this.scope = reader.ReadElementValue();
                    return true;
                case XmlElementNames.HighlightTermValue:
                    this.value = reader.ReadElementValue();
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
            if (jsonProperty.ContainsKey(XmlElementNames.HighlightTermScope))
            {
                this.scope = jsonProperty.ReadAsString(XmlElementNames.HighlightTermScope);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.HighlightTermValue))
            {
                this.value = jsonProperty.ReadAsString(XmlElementNames.HighlightTermValue);
            }
        }
    }
}
