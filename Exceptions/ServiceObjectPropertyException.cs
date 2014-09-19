// ---------------------------------------------------------------------------
// <copyright file="ServiceObjectPropertyException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceObjectPropertyException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when an operation on a property fails.
    /// </summary>
    [Serializable]
    public class ServiceObjectPropertyException : PropertyException
    {
        /// <summary>
        /// The definition of the property that is at the origin of the exception.
        /// </summary>
        private PropertyDefinitionBase propertyDefinition;

        /// <summary>
        /// ServiceObjectPropertyException constructor.
        /// </summary>
        /// <param name="propertyDefinition">The definition of the property that is at the origin of the exception.</param>
        public ServiceObjectPropertyException(PropertyDefinitionBase propertyDefinition)
            : base(propertyDefinition.GetPrintableName())
        {
            this.propertyDefinition = propertyDefinition;
        }

        /// <summary>
        /// ServiceObjectPropertyException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="propertyDefinition">The definition of the property that is at the origin of the exception.</param>
        public ServiceObjectPropertyException(string message, PropertyDefinitionBase propertyDefinition)
            : base(message, propertyDefinition.GetPrintableName())
        {
            this.propertyDefinition = propertyDefinition;
        }

        /// <summary>
        /// ServiceObjectPropertyException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="propertyDefinition">The definition of the property that is at the origin of the exception.</param>
        /// <param name="innerException">Inner exception.</param>
        public ServiceObjectPropertyException(
            string message,
            PropertyDefinitionBase propertyDefinition,
            Exception innerException)
            : base(message, propertyDefinition.GetPrintableName(), innerException)
        {
            this.propertyDefinition = propertyDefinition;
        }

        /// <summary>
        /// Gets the definition of the property that caused the exception.
        /// </summary>
        public PropertyDefinitionBase PropertyDefinition
        {
            get { return this.propertyDefinition; }
        }
    }
}
