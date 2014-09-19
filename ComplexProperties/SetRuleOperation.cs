// ---------------------------------------------------------------------------
// <copyright file="SetRuleOperation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetRuleOperation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an operation to update an existing rule.
    /// </summary>
    public sealed class SetRuleOperation : RuleOperation
    {
        /// <summary>
        /// Inbox rule to be updated.
        /// </summary>
        private Rule rule;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetRuleOperation"/> class.
        /// </summary>
        public SetRuleOperation()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SetRuleOperation"/> class.
        /// </summary>
        /// <param name="rule">The inbox rule to update.</param>
        public SetRuleOperation(Rule rule)
            : base()
        {
            this.rule = rule;
        }

        /// <summary>
        /// Gets or sets the rule to be updated.
        /// </summary>
        public Rule Rule
        {
            get
            {
                return this.rule;
            }

            set
            {
                this.SetFieldValue<Rule>(ref this.rule, value);
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
                case XmlElementNames.Rule:
                    this.rule = new Rule();
                    this.rule.LoadFromXml(reader, reader.LocalName);
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
                    case XmlElementNames.Rule:
                        this.rule = new Rule();
                        this.rule.LoadFromJson(jsonProperty.ReadAsJsonObject(key), service);
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
            this.Rule.WriteToXml(writer, XmlElementNames.Rule);
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

            jsonProperty.Add(XmlElementNames.Rule, this.Rule.InternalToJson(service));

            return jsonProperty;
        }

        /// <summary>
        ///  Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            EwsUtilities.ValidateParam(this.rule, "Rule");
        }

        /// <summary>
        /// Gets the Xml element name of the SetRuleOperation object.
        /// </summary>
        internal override string XmlElementName
        {
            get
            {
                return XmlElementNames.SetRuleOperation;
            }
        }
    }
}
