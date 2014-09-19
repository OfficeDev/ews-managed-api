// ---------------------------------------------------------------------------
// <copyright file="UpdateInboxRulesResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UpdateInboxRulesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Xml;

    /// <summary>
    /// Represents the response to a UpdateInboxRulesResponse operation.
    /// </summary>
    internal sealed class UpdateInboxRulesResponse : ServiceResponse
    {
        /// <summary>
        /// Rule operation error collection.
        /// </summary>
        private RuleOperationErrorCollection errors;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateInboxRulesResponse"/> class.
        /// </summary>
        internal UpdateInboxRulesResponse()
            : base()
        {
            this.errors = new RuleOperationErrorCollection();
        }

        /// <summary>
        /// Loads extra error details from XML
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="xmlElementName">The current element name of the extra error details.</param>
        /// <returns>True if the expected extra details is loaded; 
        /// False if the element name does not match the expected element. </returns>
        internal override bool LoadExtraErrorDetailsFromXml(EwsServiceXmlReader reader, string xmlElementName)
        {
            if (xmlElementName.Equals(XmlElementNames.MessageXml))
            {
                return base.LoadExtraErrorDetailsFromXml(reader, xmlElementName);
            }
            else if (xmlElementName.Equals(XmlElementNames.RuleOperationErrors))
            {
                this.errors.LoadFromXml(reader, XmlNamespace.Messages, xmlElementName);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the rule operation errors in the response.
        /// </summary>
        internal RuleOperationErrorCollection Errors
        {
            get
            {
                return this.errors;
            }
        }
    }
}