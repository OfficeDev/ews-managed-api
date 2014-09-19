// ---------------------------------------------------------------------------
// <copyright file="RuleError.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RuleError class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an error that occurred as a result of executing a rule operation. 
    /// </summary>
    public sealed class RuleError : ComplexProperty
    {
        /// <summary>
        /// Rule property.
        /// </summary>
        private RuleProperty ruleProperty;

        /// <summary>
        /// Rule validation error code.
        /// </summary>
        private RuleErrorCode errorCode;

        /// <summary>
        /// Error message.
        /// </summary>
        private string errorMessage;

        /// <summary>
        /// Field value.
        /// </summary>
        private string value;

        /// <summary>
        /// Initializes a new instance of the <see cref="RuleError"/> class.
        /// </summary>
        internal RuleError()
            : base()
        {
        }

        /// <summary>
        /// Gets the property which failed validation.
        /// </summary>
        public RuleProperty RuleProperty
        {
            get
            {
                return this.ruleProperty;
            }
        }

        /// <summary>
        /// Gets the validation error code.
        /// </summary>
        public RuleErrorCode ErrorCode
        {
            get
            {
                return this.errorCode;
            }
        }

        /// <summary>
        /// Gets the error message.
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return this.errorMessage;
            }
        }

        /// <summary>
        /// Gets the value that failed validation.
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
                case XmlElementNames.FieldURI:
                    this.ruleProperty = reader.ReadElementValue<RuleProperty>();
                    return true;
                case XmlElementNames.ErrorCode:
                    this.errorCode = reader.ReadElementValue<RuleErrorCode>();
                    return true;
                case XmlElementNames.ErrorMessage:
                    this.errorMessage = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FieldValue:
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
            if (jsonProperty.ContainsKey(XmlElementNames.FieldURI))
            {
                this.ruleProperty = jsonProperty.ReadEnumValue<RuleProperty>(XmlElementNames.FieldURI);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.ErrorCode))
            {
                this.errorCode = jsonProperty.ReadEnumValue<RuleErrorCode>(XmlElementNames.ErrorCode);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.ErrorMessage))
            {
                this.errorMessage = jsonProperty.ReadAsString(XmlElementNames.ErrorMessage);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.FieldValue))
            {
                this.value = jsonProperty.ReadAsString(XmlElementNames.FieldValue);
            }
        }
    }
}
