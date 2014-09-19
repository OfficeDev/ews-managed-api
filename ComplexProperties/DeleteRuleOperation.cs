// ---------------------------------------------------------------------------
// <copyright file="DeleteRuleOperation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DeleteRuleOperation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an operation to delete an existing rule.
    /// </summary>
    public sealed class DeleteRuleOperation : RuleOperation
    {
        /// <summary>
        /// Id of the inbox rule to delete.
        /// </summary>
        private string ruleId;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteRuleOperation"/> class.
        /// </summary>
        public DeleteRuleOperation()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteRuleOperation"/> class.
        /// </summary>
        /// <param name="ruleId">The Id of the inbox rule to delete.</param>
        public DeleteRuleOperation(string ruleId)
            : base()
        {
            this.ruleId = ruleId;
        }

        /// <summary>
        /// Gets or sets the Id of the rule to delete.
        /// </summary>
        public string RuleId
        {
            get
            {
                return this.ruleId;
            }

            set
            {
                this.SetFieldValue<string>(ref this.ruleId, value);
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RuleId, this.RuleId);
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

            jsonProperty.Add(XmlElementNames.RuleId, this.RuleId);

            return jsonProperty;
        }

        /// <summary>
        ///  Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            EwsUtilities.ValidateParam(this.ruleId, "RuleId");
        }

        /// <summary>
        /// Gets the Xml element name of the DeleteRuleOperation object.
        /// </summary>
        internal override string XmlElementName
        {
            get
            {
                return XmlElementNames.DeleteRuleOperation;
            }
        }
    }
}
