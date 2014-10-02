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
// <summary>Defines the MultiResponseServiceRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Represents a service request that can have multiple responses.
    /// </summary>
    /// <typeparam name="TResponse">The type of the response.</typeparam>
    internal abstract class MultiResponseServiceRequest<TResponse> : SimpleServiceRequestBase
        where TResponse : ServiceResponse
    {
        private ServiceErrorHandling errorHandlingMode;

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Service response collection.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            ServiceResponseCollection<TResponse> serviceResponses = new ServiceResponseCollection<TResponse>();

            reader.ReadStartElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

            for (int i = 0; i < this.GetExpectedResponseMessageCount(); i++)
            {
                // Read ahead to see if we've reached the end of the response messages early.
                reader.Read();
                if (reader.IsEndElement(XmlNamespace.Messages, XmlElementNames.ResponseMessages))
                {
                    break;
                }

                TResponse response = this.CreateServiceResponse(reader.Service, i);

                response.LoadFromXml(reader, this.GetResponseMessageXmlElementName());

                // Add the response to the list after it has been deserialized because the response
                // list updates an overall result as individual responses are added to it.
                serviceResponses.Add(response);
            }

            // If there's a general error in batch processing,
            // the server will return a single response message containing the error
            // (for example, if the SavedItemFolderId is bogus in a batch CreateItem
            // call). In this case, throw a ServiceResponsException. Otherwise this 
            // is an unexpected server error.
            if (serviceResponses.Count < this.GetExpectedResponseMessageCount())
            {
                if ((serviceResponses.Count == 1) && (serviceResponses[0].Result == ServiceResult.Error))
                {
                    throw new ServiceResponseException(serviceResponses[0]);
                }
                else
                {
                    throw new ServiceXmlDeserializationException(
                                string.Format(
                                    Strings.TooFewServiceReponsesReturned,
                                    this.GetResponseMessageXmlElementName(),
                                    this.GetExpectedResponseMessageCount(),
                                    serviceResponses.Count));
                }
            }

            reader.ReadEndElementIfNecessary(XmlNamespace.Messages, XmlElementNames.ResponseMessages);

            return serviceResponses;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="jsonBody">The json body.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(JsonObject jsonBody)
        {
            ServiceResponseCollection<TResponse> serviceResponses = new ServiceResponseCollection<TResponse>();

            object[] jsonResponseMessages = jsonBody.ReadAsJsonObject(XmlElementNames.ResponseMessages).ReadAsArray(XmlElementNames.Items);

            int responseCtr = 0;
            foreach (object jsonResponseObject in jsonResponseMessages)
            {
                TResponse response = this.CreateServiceResponse(this.Service, responseCtr);

                response.LoadFromJson(jsonResponseObject as JsonObject, this.Service);

                // Add the response to the list after it has been deserialized because the response
                // list updates an overall result as individual responses are added to it.
                serviceResponses.Add(response);

                responseCtr++;
            }

            if (serviceResponses.Count < this.GetExpectedResponseMessageCount())
            {
                if ((serviceResponses.Count == 1) && (serviceResponses[0].Result == ServiceResult.Error))
                {
                    throw new ServiceResponseException(serviceResponses[0]);
                }
                else
                {
                    throw new ServiceJsonDeserializationException();
                }
            }

            return serviceResponses;
        }

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="responseIndex">Index of the response.</param>
        /// <returns>Service response.</returns>
        internal abstract TResponse CreateServiceResponse(ExchangeService service, int responseIndex);
        
        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetResponseMessageXmlElementName();

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal abstract int GetExpectedResponseMessageCount();

        /// <summary>
        /// Initializes a new instance of the <see cref="MultiResponseServiceRequest&lt;TResponse&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="errorHandlingMode"> Indicates how errors should be handled.</param>
        internal MultiResponseServiceRequest(ExchangeService service, ServiceErrorHandling errorHandlingMode)
            : base(service)
        {
            this.errorHandlingMode = errorHandlingMode;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response collection.</returns>
        internal ServiceResponseCollection<TResponse> Execute()
        {
            ServiceResponseCollection<TResponse> serviceResponses = (ServiceResponseCollection<TResponse>)this.InternalExecute();

            if (this.ErrorHandlingMode == ServiceErrorHandling.ThrowOnError)
            {
                EwsUtilities.Assert(
                    serviceResponses.Count == 1,
                    "MultiResponseServiceRequest.Execute",
                    "ServiceErrorHandling.ThrowOnError error handling is only valid for singleton request");

                serviceResponses[0].ThrowIfNecessary();
            }

            return serviceResponses;
        }

        /// <summary>
        /// Ends executing this async request.
        /// </summary>
        /// <param name="asyncResult">The async result</param>
        /// <returns>Service response collection.</returns>
        internal ServiceResponseCollection<TResponse> EndExecute(IAsyncResult asyncResult)
        {
            ServiceResponseCollection<TResponse> serviceResponses = (ServiceResponseCollection<TResponse>)this.EndInternalExecute(asyncResult);

            if (this.ErrorHandlingMode == ServiceErrorHandling.ThrowOnError)
            {
                EwsUtilities.Assert(
                    serviceResponses.Count == 1,
                    "MultiResponseServiceRequest.Execute",
                    "ServiceErrorHandling.ThrowOnError error handling is only valid for singleton request");

                serviceResponses[0].ThrowIfNecessary();
            }

            return serviceResponses;
        }

        /// <summary>
        /// Gets a value indicating how errors should be handled.
        /// </summary>
        internal ServiceErrorHandling ErrorHandlingMode
        {
            get { return this.errorHandlingMode; }
        }
    }
}
