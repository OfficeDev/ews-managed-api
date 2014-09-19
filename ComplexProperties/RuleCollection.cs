// ---------------------------------------------------------------------------
// <copyright file="RuleCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements a rule collection.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Xml;

    /// <summary>
    /// Represents a collection of rules.
    /// </summary>
    public sealed class RuleCollection : ComplexProperty, IEnumerable<Rule>, IJsonCollectionDeserializer
    {
        /// <summary>
        /// The OutlookRuleBlobExists flag.
        /// </summary>
        private bool outlookRuleBlobExists;

        /// <summary>
        /// The rules in the rule collection.
        /// </summary>
        private List<Rule> rules;

        /// <summary>
        /// Initializes a new instance of the <see cref="RuleCollection"/> class.
        /// </summary>
        internal RuleCollection()
            : base()
        {
            this.rules = new List<Rule>();
        }

        /// <summary>
        /// Gets a value indicating whether an Outlook rule blob exists in the user's
        /// mailbox. To update rules with EWS when the Outlook rule blob exists, call
        /// SetInboxRules passing true as the value of the removeOutlookBlob parameter.
        /// </summary>
        public bool OutlookRuleBlobExists
        {
            get
            {
                return this.outlookRuleBlobExists;
            }

            internal set
            {
                this.outlookRuleBlobExists = value;
            }
        }

        /// <summary>
        /// Gets the number of rules in this collection.
        /// </summary>
        public int Count
        {
            get
            {
                return this.rules.Count;
            }
        }

        /// <summary>
        /// Gets the rule at the specified index in the collection.
        /// </summary>
        /// <param name="index">The index of the rule to get.</param>
        /// <returns>The rule at the specified index.</returns>
        public Rule this[int index]
        {
            get
            {
                if (index < 0 || index >= this.rules.Count)
                {
                    throw new ArgumentOutOfRangeException("Index");
                }

                return this.rules[index];
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.Rule))
            {
                Rule rule = new Rule();
                rule.LoadFromXml(reader, XmlElementNames.Rule);
                this.rules.Add(rule);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Loads from json collection.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.CreateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            foreach (object entry in jsonCollection)
            {
                JsonObject jsonEntry = entry as JsonObject;
                if (jsonEntry != null)
                {
                    Rule rule = new Rule();
                    rule.LoadFromJson(jsonEntry, service);
                    this.rules.Add(rule);
                }
            }
        }

        /// <summary>
        /// Loads from json collection to update the existing collection element.
        /// </summary>
        /// <param name="jsonCollection">The json collection.</param>
        /// <param name="service">The service.</param>
        void IJsonCollectionDeserializer.UpdateFromJsonCollection(object[] jsonCollection, ExchangeService service)
        {
            throw new NotImplementedException();
        }

        #region IEnumerable Interface
        /// <summary>
        /// Get an enumerator for the collection
        /// </summary>
        /// <returns>Enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        /// <summary>
        /// Get an enumerator for the collection
        /// </summary>
        /// <returns>Enumerator</returns>
        public IEnumerator<Rule> GetEnumerator()
        {
            return rules.GetEnumerator();
        }
        #endregion
    }
}
