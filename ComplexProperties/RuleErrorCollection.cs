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
// <summary>Implements a RuleErrorCollection collection.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a collection of rule validation errors.
    /// </summary>
    internal sealed class RuleErrorCollection : ComplexPropertyCollection<RuleError>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RuleErrorCollection"/> class.
        /// </summary>
        internal RuleErrorCollection()
            : base()
        {
        }

        /// <summary>
        /// Creates an RuleError object from an XML element name.
        /// </summary>
        /// <param name="xmlElementName">The XML element name from which to create the RuleError object.</param>
        /// <returns>A RuleError object.</returns>
        internal override RuleError CreateComplexProperty(string xmlElementName)
        {
            if (xmlElementName == XmlElementNames.Error)
            {
                return new RuleError();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Creates the default complex property.
        /// </summary>
        /// <returns>A RuleError object.</returns>
        internal override RuleError CreateDefaultComplexProperty()
        {
            return new RuleError();
        }

        /// <summary>
        /// Retrieves the XML element name corresponding to the provided RuleError object.
        /// </summary>
        /// <param name="ruleValidationError">The RuleError object from which to determine the XML element name.</param>
        /// <returns>The XML element name corresponding to the provided RuleError object.</returns>
        internal override string GetCollectionItemXmlElementName(RuleError ruleValidationError)
        {
            return XmlElementNames.Error;
        }
    }
}
