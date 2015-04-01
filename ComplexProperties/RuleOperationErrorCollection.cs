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
    /// Represents a collection of rule operation errors.
    /// </summary>
    public sealed class RuleOperationErrorCollection : ComplexPropertyCollection<RuleOperationError>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RuleOperationErrorCollection"/> class.
        /// </summary>
        internal RuleOperationErrorCollection()
            : base()
        {
        }

        /// <summary>
        /// Creates an RuleOperationError object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to create the RuleOperationError object.</param>
        /// <returns>A RuleOperationError object.</returns>
        internal override RuleOperationError CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == XmlElementNames.RuleOperationError)
            {
                return new RuleOperationError();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>A RuleOperationError object.</returns>
        internal override RuleOperationError CreateDefaultComplexProperty()
        {
            return new RuleOperationError();
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided RuleOperationError object.
        /// </summary>
        /// <param name="operationError">The RuleOperationError object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided RuleOperationError object.</returns>
        internal override string GetCollectionItemXmlElementName(RuleOperationError operationError)
        {
            return XmlElementNames.RuleOperationError;
        }
    }
}