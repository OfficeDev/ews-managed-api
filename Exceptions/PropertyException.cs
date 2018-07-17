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
    public class PropertyException : ServiceLocalException
    {
        /// <summary>
        /// The name of the property that is at the origin of the exception.
        /// </summary>
        private readonly string name;

        /// <summary>
        /// PropertyException constructor.
        /// </summary>
        /// <param name="name">The name of the property that is at the origin of the exception.</param>
        public PropertyException(string name)
            : base()
        {
            this.name = name;
        }

        /// <summary>
        /// PropertyException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="name">The name of the property that is at the origin of the exception.</param>
        public PropertyException(string message, string name)
            : base(message)
        {
            this.name = name;
        }

        /// <summary>
        /// PropertyException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="name">The name of the property that is at the origin of the exception.</param>
        /// <param name="innerException">Inner exception.</param>
        public PropertyException(
            string message,
            string name,
            Exception innerException)
            : base(message, innerException)
        {
            this.name = name;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="T:Microsoft.Exchange.WebServices.Data.PropertyException"/> class with serialized data.
		/// </summary>
		/// <param name="info">The object that holds the serialized object data.</param>
		/// <param name="context">The contextual information about the source or destination.</param>
		protected PropertyException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
			this.name = info.GetString("PropertyName");
		}

		/// <summary>Sets the <see cref="T:System.Runtime.Serialization.SerializationInfo" /> object with the parameter name and additional exception information.</summary>
		/// <param name="info">The object that holds the serialized object data. </param>
		/// <param name="context">The contextual information about the source or destination. </param>
		/// <exception cref="T:System.ArgumentNullException">The <paramref name="info" /> object is a null reference (Nothing in Visual Basic). </exception>
		public override void GetObjectData(SerializationInfo info, StreamingContext context)
		{
			EwsUtilities.Assert(info != null, "PropertyException.GetObjectData", "info is null");

			base.GetObjectData(info, context);

			info.AddValue("PropertyName", this.name);
		}

		/// <summary>
		/// Gets the name of the property that caused the exception.
		/// </summary>
		public string Name
        {
            get { return this.name; }
        }
    }
}