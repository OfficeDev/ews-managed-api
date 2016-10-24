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
    using System;
    using System.Runtime.Serialization;

    /// <summary>
    /// Represents an error that occurs when an operation on a property fails.
    /// </summary>
    [Serializable]
    public class ServiceObjectPropertyException : PropertyException
    {
        /// <summary>
        /// The definition of the property that is at the origin of the exception.
        /// </summary>
        private readonly PropertyDefinitionBase propertyDefinition;

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
		/// Initializes a new instance of the <see cref="T:Microsoft.Exchange.WebServices.Data.ServiceObjectPropertyException"/> class with serialized data.
		/// </summary>
		/// <param name="info">The object that holds the serialized object data.</param>
		/// <param name="context">The contextual information about the source or destination.</param>
		protected ServiceObjectPropertyException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
			this.propertyDefinition = (PropertyDefinitionBase)info.GetValue("PropertyDefinition", typeof(PropertyDefinitionBase));
		}

		/// <summary>Sets the <see cref="T:System.Runtime.Serialization.SerializationInfo" /> object with the parameter name and additional exception information.</summary>
		/// <param name="info">The object that holds the serialized object data. </param>
		/// <param name="context">The contextual information about the source or destination. </param>
		/// <exception cref="T:System.ArgumentNullException">The <paramref name="info" /> object is a null reference (Nothing in Visual Basic). </exception>
		public override void GetObjectData(SerializationInfo info, StreamingContext context)
		{
			EwsUtilities.Assert(info != null, "ServiceObjectPropertyException.GetObjectData", "info is null");

			base.GetObjectData(info, context);

			info.AddValue("PropertyDefinition", this.propertyDefinition, typeof(PropertyDefinitionBase));
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