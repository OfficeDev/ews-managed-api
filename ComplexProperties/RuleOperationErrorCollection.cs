// ---------------------------------------------------------------------------
// <copyright file="RuleOperationErrorCollection.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Implements a RuleOperationErrorCollection collection.</summary>
//-----------------------------------------------------------------------

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
