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
    /// Represents a remote service exception that can have multiple service responses.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    [Serializable]
    public abstract class BatchServiceResponseException<TResponse> : ServiceRemoteException
        where TResponse : ServiceResponse
    {
        /// <summary>
        /// The list of responses returned by the web method.
        /// </summary>
        private readonly ServiceResponseCollection<TResponse> responses;

        /// <summary>
        /// Initializes a new instance of MultiServiceResponseException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        internal BatchServiceResponseException(
            ServiceResponseCollection<TResponse> serviceResponses,
            string message)
            : base(message)
        {
            EwsUtilities.Assert(
                serviceResponses != null,
                "MultiServiceResponseException.ctor",
                "serviceResponses is null");

            this.responses = serviceResponses;
        }

        /// <summary>
        /// Initializes a new instance of MultiServiceResponseException.
        /// </summary>
        /// <param name="serviceResponses">The list of responses to be associated with this exception.</param>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        internal BatchServiceResponseException(
            ServiceResponseCollection<TResponse> serviceResponses,
            string message,
            Exception innerException)
            : base(message, innerException)
        {
            EwsUtilities.Assert(
                serviceResponses != null,
                "MultiServiceResponseException.ctor",
                "serviceResponses is null");

            this.responses = serviceResponses;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="T:Microsoft.Exchange.WebServices.Data.BatchServiceResponseException"/> class with serialized data.
		/// </summary>
		/// <param name="info">The object that holds the serialized object data.</param>
		/// <param name="context">The contextual information about the source or destination.</param>
		protected BatchServiceResponseException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
			this.responses = (ServiceResponseCollection<TResponse>)info.GetValue("Responses", typeof(ServiceResponseCollection<TResponse>));
		}

		/// <summary>Sets the <see cref="T:System.Runtime.Serialization.SerializationInfo" /> object with the parameter name and additional exception information.</summary>
		/// <param name="info">The object that holds the serialized object data. </param>
		/// <param name="context">The contextual information about the source or destination. </param>
		/// <exception cref="T:System.ArgumentNullException">The <paramref name="info" /> object is a null reference (Nothing in Visual Basic). </exception>
		public override void GetObjectData(SerializationInfo info, StreamingContext context)
		{
			EwsUtilities.Assert(info != null, "BatchServiceResponseException.GetObjectData", "info is null");

			base.GetObjectData(info, context);

			info.AddValue("Responses", this.responses, typeof(ServiceResponseCollection<TResponse>));
		}

		/// <summary>
		/// Gets a list of responses returned by the web method.
		/// </summary>
		public ServiceResponseCollection<TResponse> ServiceResponses
        {
            get { return this.responses; }
        }
    }
}