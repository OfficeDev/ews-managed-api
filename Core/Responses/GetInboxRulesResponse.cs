// ---------------------------------------------------------------------------
// <copyright file="GetInboxRulesResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetInboxRulesResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a GetInboxRules operation.
    /// </summary>
    internal sealed class GetInboxRulesResponse : ServiceResponse
    {
        /// <summary>
        /// Rule collection.
        /// </summary>
        private RuleCollection ruleCollection;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetInboxRulesResponse"/> class.
        /// </summary>
        internal GetInboxRulesResponse()
            : base()
        {
            this.ruleCollection = new RuleCollection();
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            reader.Read();
            this.ruleCollection.OutlookRuleBlobExists = reader.ReadElementValue<bool>(
                XmlNamespace.Messages, 
                XmlElementNames.OutlookRuleBlobExists);
            reader.Read();
            if (reader.IsStartElement(XmlNamespace.NotSpecified, XmlElementNames.InboxRules))
            {
                this.ruleCollection.LoadFromXml(reader, XmlNamespace.NotSpecified, XmlElementNames.InboxRules);
            }
        }

        /// <summary>
        /// Gets the rule collection in the response.
        /// </summary>
        internal RuleCollection Rules
        {
            get
            {
                return this.ruleCollection;
            }
        }
    }
}