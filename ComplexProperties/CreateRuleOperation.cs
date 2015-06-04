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
    /// <summary>
    /// Represents an operation to create a new rule.
    /// </summary>
    public sealed class CreateRuleOperation : RuleOperation
    {
        /// <summary>
        /// Inbox rule to be created.
        /// </summary>
        private Rule rule;

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateRuleOperation"/> class.
        /// </summary>
        public CreateRuleOperation()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateRuleOperation"/> class.
        /// </summary>
        /// <param name="rule">The inbox rule to create.</param>
        public CreateRuleOperation(Rule rule)
            : base()
        {
            this.rule = rule;
        }

        /// <summary>
        /// Gets or sets the rule to be created.
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
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            this.Rule.WriteToXml(writer, XmlElementNames.Rule);
        }

        /// <summary>
        ///  Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            EwsUtilities.ValidateParam(this.rule, "Rule");
        }

        /// <summary>
        /// Gets the Xml element name of the CreateRuleOperation object.
        /// </summary>
        internal override string XmlElementName
        {
            get
            {
                return XmlElementNames.CreateRuleOperation;
            }
        }
    }
}