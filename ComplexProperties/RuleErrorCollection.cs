// ---------------------------------------------------------------------------
// <copyright file="RuleErrorCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
