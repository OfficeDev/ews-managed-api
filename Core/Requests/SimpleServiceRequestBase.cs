// ---------------------------------------------------------------------------
// <copyright file="SimpleServiceRequestBase.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SimpleServiceRequestBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;
    using System.IO.Compression;
    using System.Net;
    using System.Xml;

    /// <summary>
    /// Represents an abstract, simple request-response service request.
    /// </summary>
    internal abstract class SimpleServiceRequestBase : ServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SimpleServiceRequestBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SimpleServiceRequestBase(ExchangeService service) :
            base(service)
        {
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal object InternalExecute()
        {
            IEwsHttpWebRequest request;
            IEwsHttpWebResponse response = this.ValidateAndEmitRequest(out request);

            return this.ReadResponse(response);
        }

        /// <summary>
        /// Ends executing this async request.
        /// </summary>
        /// <param name="asyncResult">The async result</param>
        /// <returns>Service response object.</returns>
        internal object EndInternalExecute(IAsyncResult asyncResult)
        {
            // We have done enough validation before
            AsyncRequestResult asyncRequestResult = (AsyncRequestResult)asyncResult;

            IEwsHttpWebResponse response = this.EndGetEwsHttpWebResponse(asyncRequestResult.WebRequest, asyncRequestResult.WebAsyncResult);
            return this.ReadResponse(response);
        }

        /// <summary>
        /// Begins executing this async request.
        /// </summary>
        /// <param name="callback">The AsyncCallback delegate.</param>
        /// <param name="state">An object that contains state information for this request.</param>
        /// <returns>An IAsyncResult that references the asynchronous request.</returns>
        internal IAsyncResult BeginExecute(AsyncCallback callback, object state)
        {
            this.Validate();

            IEwsHttpWebRequest request = this.BuildEwsHttpWebRequest();

            WebAsyncCallStateAnchor wrappedState = new WebAsyncCallStateAnchor(this, request, callback /* user callback */, state /* user state */);

            // BeginGetResponse() does not throw interesting exceptions
            IAsyncResult webAsyncResult = request.BeginGetResponse(SimpleServiceRequestBase.WebRequestAsyncCallback, wrappedState);

            return new AsyncRequestResult(this, request, webAsyncResult, state /* user state */);
        }

        /// <summary>
        /// Async callback method for HttpWebRequest async requests.
        /// </summary>
        /// <param name="webAsyncResult">An IAsyncResult that references the asynchronous request.</param>
        private static void WebRequestAsyncCallback(IAsyncResult webAsyncResult)
        {
            WebAsyncCallStateAnchor wrappedState = webAsyncResult.AsyncState as WebAsyncCallStateAnchor;

            if (wrappedState != null && wrappedState.AsyncCallback != null)
            {
                AsyncRequestResult asyncRequestResult = new AsyncRequestResult(
                    wrappedState.ServiceRequest,
                    wrappedState.WebRequest,
                    webAsyncResult, /* web async result */
                    wrappedState.AsyncState /* user state */);

                // Call user's call back
                wrappedState.AsyncCallback(asyncRequestResult);
            }
        }

        /// <summary>
        /// Reads the response with error handling
        /// </summary>
        /// <param name="response">The response.</param>
        /// <returns>Service response.</returns>
        private object ReadResponse(IEwsHttpWebResponse response)
        {
            object serviceResponse;

            try
            {
                this.Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, response);

                // If tracing is enabled, we read the entire response into a MemoryStream so that we
                // can pass it along to the ITraceListener. Then we parse the response from the 
                // MemoryStream.
                if (this.Service.IsTraceEnabledFor(TraceFlags.EwsResponse))
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        using (Stream serviceResponseStream = ServiceRequestBase.GetResponseStream(response))
                        {
                            // Copy response to in-memory stream and reset position to start.
                            EwsUtilities.CopyStream(serviceResponseStream, memoryStream);
                            memoryStream.Position = 0;
                        }

                        if (this.Service.RenderingMethod == ExchangeService.RenderingMode.Xml)
                        {
                            this.TraceResponseXml(response, memoryStream);

                            serviceResponse = this.ReadResponseXml(memoryStream);
                        }
                        else if (this.Service.RenderingMethod == ExchangeService.RenderingMode.JSON)
                        {
                            this.TraceResponseJson(response, memoryStream);

                            serviceResponse = this.ReadResponseJson(memoryStream);
                        }
                        else
                        {
                            throw new InvalidOperationException("Unknown RenderingMethod.");
                        }
                    }
                }
                else
                {
                    using (Stream responseStream = ServiceRequestBase.GetResponseStream(response))
                    {
                        if (this.Service.RenderingMethod == ExchangeService.RenderingMode.Xml)
                        {
                            serviceResponse = this.ReadResponseXml(responseStream);
                        }
                        else if (this.Service.RenderingMethod == ExchangeService.RenderingMode.JSON)
                        {
                            serviceResponse = this.ReadResponseJson(responseStream);
                        }
                        else
                        {
                            throw new InvalidOperationException("Unknown RenderingMethod.");
                        }
                    }
                }
            }
            catch (WebException e)
            {
                if (e.Response != null)
                {
                    IEwsHttpWebResponse exceptionResponse = this.Service.HttpWebRequestFactory.CreateExceptionResponse(e);
                    this.Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, exceptionResponse);
                }

                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
            catch (IOException e)
            {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }

            return serviceResponse;
        }

        /// <summary>
        /// Reads the response json.
        /// </summary>
        /// <param name="responseStream">The response stream.</param>
        /// <returns></returns>
        private object ReadResponseJson(Stream responseStream)
        {
            JsonObject jsonResponse = new JsonParser(responseStream).Parse();
            return this.BuildResponseObjectFromJson(jsonResponse);
        }

        /// <summary>
        /// Reads the response XML.
        /// </summary>
        /// <param name="responseStream">The response stream.</param>
        /// <returns></returns>
        private object ReadResponseXml(Stream responseStream)
        {
            object serviceResponse;
            EwsServiceXmlReader ewsXmlReader = new EwsServiceXmlReader(responseStream, this.Service);
            serviceResponse = this.ReadResponse(ewsXmlReader);
            return serviceResponse;
        }
    }
}