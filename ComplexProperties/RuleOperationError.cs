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
// <summary>Defines the RuleOperationError class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents an error that occurred while processing a rule operation.
    /// </summary>
    public sealed class RuleOperationError : ComplexProperty, IEnumerable<RuleError>
    {
        /// <summary>
        /// Index of the operation mapping to the error.
        /// </summary>
        private int operationIndex;

        /// <summary>
        /// RuleOperation object mapping to the error.
        /// </summary>
        private RuleOperation operation;

        /// <summary>
        /// RuleError Collection.
        /// </summary>
        private RuleErrorCollection ruleErrors;

        /// <summary>
        /// Initializes a new instance of the <see cref="RuleOperationError"/> class.
        /// </summary>
        internal RuleOperationError()
            : base()
        {
        }

        /// <summary>
        /// Gets the operation that resulted in an error.
        /// </summary>
        public RuleOperation Operation
        {
            get { return this.operation; }
        }

        /// <summary>
        /// Gets the number of rule errors in the list.
        /// </summary>
        public int Count
        {
            get { return this.ruleErrors.Count; }
        }

        /// <summary>
        /// Gets the rule error at the specified index.
        /// </summary>
        /// <param name="index">The index of the rule error to get.</param>
        /// <returns>The rule error at the specified index.</returns>
        public RuleError this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    throw new ArgumentOutOfRangeException("index");
                }

                return this.ruleErrors[index];
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
                case XmlElementNames.OperationIndex:
                    this.operationIndex = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.ValidationErrors:
                    this.ruleErrors = new RuleErrorCollection();
                    this.ruleErrors.LoadFromXml(reader, reader.LocalName);
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
            if (jsonProperty.ContainsKey(XmlElementNames.OperationIndex))
            {
                this.operationIndex = jsonProperty.ReadAsInt(XmlElementNames.OperationIndex);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.ValidationErrors))
            {
                this.ruleErrors = new RuleErrorCollection();
                (this.ruleErrors as IJsonCollectionDeserializer).CreateFromJsonCollection(jsonProperty.ReadAsArray(XmlElementNames.ValidationErrors), service);
            }
        }

        /// <summary>
        /// Set operation property by the index of a given opeation enumerator.
        /// </summary>
        /// <param name="operations">Operation enumerator.</param>
        internal void SetOperationByIndex(IEnumerator<RuleOperation> operations)
        {
            operations.Reset();
            for (int i = 0; i <= this.operationIndex; i++)
            {
                operations.MoveNext();
            }
            this.operation = operations.Current;
        }

        #region IEnumerable<RuleError> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<RuleError> GetEnumerator()
        {
            return this.ruleErrors.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.ruleErrors.GetEnumerator();
        }

        #endregion
    }
}
