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
