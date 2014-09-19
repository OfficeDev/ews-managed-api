// ---------------------------------------------------------------------------
// <copyright file="RuleOperation.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the RuleOperation class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an operation to be performed on a rule.
    /// </summary>
    public abstract class RuleOperation : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RuleOperation"/> class.
        /// </summary>
        internal RuleOperation()
            : base()
        {
        }

        /// <summary>
        /// Gets the XML element name of the rule operation.
        /// </summary>
        internal abstract string XmlElementName
        {
            get;
        }
    }
}
