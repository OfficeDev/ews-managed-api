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