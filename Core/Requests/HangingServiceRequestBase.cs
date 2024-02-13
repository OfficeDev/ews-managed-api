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
    using System.IO;
    using System.IO.Compression;
    using System.Net;
    using System.Text;
    using System.Threading;
    using System.Web;
    using System.Xml;

    /// <summary>
    /// Enumeration of reasons that a hanging request may disconnect.
    /// </summary>
    internal enum HangingRequestDisconnectReason
    {
        /// <summary>The server cleanly closed the connection.</summary>
        Clean,

        /// <summary>The client closed the connection.</summary>
        UserInitiated,

        /// <summary>The connection timed out do to a lack of a heartbeat received.</summary>
        Timeout,

        /// <summary>An exception occurred on the connection</summary>
        Exception
    }

    /// <summary>
    /// Represents a collection of arguments for the HangingServiceRequestBase.HangingRequestDisconnectHandler
    /// delegate method.
    /// </summary>
    internal class HangingRequestDisconnectEventArgs : EventArgs
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HangingRequestDisconnectEventArgs"/> class.
        /// </summary>
        /// <param name="reason">The reason.</param>
        /// <param name="exception">The exception.</param>
        internal HangingRequestDisconnectEventArgs(
            HangingRequestDisconnectReason reason,
            Exception exception)
        {
            this.Reason = reason;
            this.Exception = exception;
        }

        /// <summary>
        /// Gets the reason that the user was disconnected.
        /// </summary>
        public HangingRequestDisconnectReason Reason
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the exception that caused the disconnection. Can be null.
        /// </summary>
        public Exception Exception
        {
            get;
            internal set;
        }
    }

    /// <summary>
    /// Represents an abstract, hanging service request.
    /// </summary>
    internal abstract class HangingServiceRequestBase : ServiceRequestBase
    {
        /// <summary>
        /// Callback delegate to handle asynchronous responses.
        /// </summary>
        /// <param name="response">Response received from the server</param>
        internal delegate void HandleResponseObject(object response);

        private const int BufferSize = 4096;

        /// <summary>
        /// Test switch to log all bytes that come across the wire.
        /// Helpful when parsing fails before certain bytes hit the trace logs.
        /// </summary>
        internal static bool LogAllWireBytes = false;

        /// <summary>
        /// Callback delegate to handle response objects
        /// </summary>
        private HandleResponseObject responseHandler;

        /// <summary>
        /// Response from the server.
        /// </summary>
        private IEwsHttpWebResponse response;

        /// <summary>
        /// Request to the server.
        /// </summary>
        private IEwsHttpWebRequest request;

        /// <summary>
        /// Expected minimum frequency in responses, in milliseconds.
        /// </summary>
        protected int heartbeatFrequencyMilliseconds;

        /// <summary>
        /// lock object
        /// </summary>
        private object lockObject = new object();

        /// <summary>
        /// Delegate method to handle a hanging request disconnection.
        /// </summary>
        /// <param name="sender">The object invoking the delegate.</param>
        /// <param name="args">Event data.</param>
        internal delegate void HangingRequestDisconnectHandler(object sender, HangingRequestDisconnectEventArgs args);

        /// <summary>
        /// Occurs when the hanging request is disconnected.
        /// </summary>
        internal event HangingRequestDisconnectHandler OnDisconnect;

        /// <summary>
        /// Initializes a new instance of the <see cref="HangingServiceRequestBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="handler">Callback delegate to handle response objects</param>
        /// <param name="heartbeatFrequency">Frequency at which we expect heartbeats, in milliseconds.</param>
        internal HangingServiceRequestBase(ExchangeService service, HandleResponseObject handler, int heartbeatFrequency) :
            base(service)
        {
            this.responseHandler = handler;
            this.heartbeatFrequencyMilliseconds = heartbeatFrequency;
        }

        /// <summary>
        /// Exectures the request.
        /// </summary>
        internal void InternalExecute()
        {
            lock (this.lockObject)
            {
                this.response = this.ValidateAndEmitRequest(out this.request);

                this.InternalOnConnect();
            }
        }

        /// <summary>
        /// Parses the responses.
        /// </summary>
        /// <param name="state">The state.</param>
        private void ParseResponses(object state)
        {
            try
            {
                Guid traceId = Guid.Empty;
                HangingTraceStream tracingStream = null;
                MemoryStream responseCopy = null;

                try
                {
                    bool traceEwsResponse = this.Service.IsTraceEnabledFor(TraceFlags.EwsResponse);

                    using (Stream responseStream = this.response.GetResponseStream())
                    {
                        responseStream.ReadTimeout = 2 * this.heartbeatFrequencyMilliseconds;
                        tracingStream = new HangingTraceStream(responseStream, this.Service);

                        // EwsServiceMultiResponseXmlReader.Create causes a read.
                        if (traceEwsResponse)
                        {
                            responseCopy = new MemoryStream();
                            tracingStream.SetResponseCopy(responseCopy);
                        }

                        EwsServiceMultiResponseXmlReader ewsXmlReader = EwsServiceMultiResponseXmlReader.Create(tracingStream, this.Service);

                        while (this.IsConnected)
                        {
                            object responseObject = null;
                            if (traceEwsResponse)
                            {
                                try
                                {
                                    responseObject = this.ReadResponse(ewsXmlReader, this.response.Headers);
                                }
                                finally
                                {
                                    this.Service.TraceXml(TraceFlags.EwsResponse, responseCopy);
                                }

                                // reset the stream collector.
                                responseCopy.Close();
                                responseCopy = new MemoryStream();
                                tracingStream.SetResponseCopy(responseCopy);
                            }
                            else
                            {
                                responseObject = this.ReadResponse(ewsXmlReader, this.response.Headers);
                            }

                            this.responseHandler(responseObject);
                        }
                    }
                }
                catch (TimeoutException ex)
                {
                    // The connection timed out.
                    this.Disconnect(HangingRequestDisconnectReason.Timeout, ex);
                    return;
                }
                catch (IOException ex)
                {
                    // Stream is closed, so disconnect.
                    this.Disconnect(HangingRequestDisconnectReason.Exception, ex);
                    return;
                }
                //catch (HttpException ex)
                //{
                //    // Stream is closed, so disconnect.
                //    this.Disconnect(HangingRequestDisconnectReason.Exception, ex);
                //    return;
                //}
                catch (WebException ex)
                {
                    // Stream is closed, so disconnect.
                    this.Disconnect(HangingRequestDisconnectReason.Exception, ex);
                    return;
                }
                catch (ObjectDisposedException ex)
                {
                    // Stream is closed, so disconnect.
                    this.Disconnect(HangingRequestDisconnectReason.Exception, ex);
                    return;
                }
                catch (NotSupportedException)
                {
                    // This is thrown if we close the stream during a read operation due to a user method call.
                    // Trying to delay closing until the read finishes simply results in a long-running connection.
                    this.Disconnect(HangingRequestDisconnectReason.UserInitiated, null);
                    return;
                }
                catch (XmlException ex)
                {
                    // Thrown if server returned no XML document.
                    this.Disconnect(HangingRequestDisconnectReason.UserInitiated, ex);
                    return;
                }
                finally
                {
                    if (responseCopy != null)
                    {
                        responseCopy.Dispose();
                        responseCopy = null;
                    }
                }
            }
            catch (ServiceLocalException exception)
            {
                this.Disconnect(HangingRequestDisconnectReason.Exception, exception);
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is connected.
        /// </summary>
        /// <value><c>true</c> if this instance is connected; otherwise, <c>false</c>.</value>
        internal bool IsConnected
        {
            get;
            private set;
        }

        /// <summary>
        /// Disconnects the request.
        /// </summary>
        internal void Disconnect()
        {
            lock (this.lockObject)
            {
                this.request.Abort();
                this.response.Close();
                this.Disconnect(HangingRequestDisconnectReason.UserInitiated, null);
            }
        }

        /// <summary>
        /// Disconnects the request with the specified reason and exception.
        /// </summary>
        /// <param name="reason">The reason.</param>
        /// <param name="exception">The exception.</param>
        internal void Disconnect(HangingRequestDisconnectReason reason, Exception exception)
        {
            if (this.IsConnected)
            {
                this.response.Close();
                this.InternalOnDisconnect(reason, exception);
            }
        }

        /// <summary>
        /// Perform any bookkeeping needed when we connect 
        /// </summary>
        private void InternalOnConnect()
        {
            if (!this.IsConnected)
            {
                this.IsConnected = true;

                // Trace Http headers
                this.Service.ProcessHttpResponseHeaders(
                    TraceFlags.EwsResponseHttpHeaders,
                    this.response);

                ThreadPool.QueueUserWorkItem(
                    new WaitCallback(this.ParseResponses));
            }
        }

        /// <summary>
        /// Perform any bookkeeping needed when we disconnect (cleanly or forcefully)
        /// </summary>
        /// <param name="reason"></param>
        /// <param name="exception"></param>
        private void InternalOnDisconnect(HangingRequestDisconnectReason reason, Exception exception)
        {
            if (this.IsConnected)
            {
                this.IsConnected = false;

                this.OnDisconnect(
                    this,
                    new HangingRequestDisconnectEventArgs(reason, exception));
            }
        }

        /// <summary>
        /// Reads any preamble data not part of the core response.
        /// </summary>
        /// <param name="ewsXmlReader">The EwsServiceXmlReader.</param>
        protected override void ReadPreamble(EwsServiceXmlReader ewsXmlReader)
        {
            // Do nothing.
        }
    }
}