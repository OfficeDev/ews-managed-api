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
// <summary>Defines the ServiceRequestBase class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Net;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents an abstract service request.
    /// </summary>
    internal abstract class ServiceRequestBase
    {
        #region Private Constants

        private static readonly string[] RequestIdResponseHeaders = new[] { "RequestId", "request-id", };
        private const string XMLSchemaNamespace = "http://www.w3.org/2001/XMLSchema";
        private const string XMLSchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";
        private const string ClientStatisticsRequestHeader = "X-ClientStatistics";

        #endregion

        /// <summary>
        /// Maintains the collection of client side statistics for requests already completed
        /// </summary>
        private static List<string> clientStatisticsCache = new List<string>();

        private ExchangeService service;

        /// <summary>
        /// Gets the response stream (may be wrapped with GZip/Deflate stream to decompress content)
        /// </summary>
        /// <param name="response">HttpWebResponse.</param>
        /// <returns>ResponseStream</returns>
        protected static Stream GetResponseStream(IEwsHttpWebResponse response)
        {
            string contentEncoding = response.ContentEncoding;
            Stream responseStream = response.GetResponseStream();

            return WrapStream(responseStream, response.ContentEncoding);
        }

        /// <summary>
        /// Gets the response stream (may be wrapped with GZip/Deflate stream to decompress content)
        /// </summary>
        /// <param name="response">HttpWebResponse.</param>
        /// <param name="readTimeout">read timeout in milliseconds</param>
        /// <returns>ResponseStream</returns>
        protected static Stream GetResponseStream(IEwsHttpWebResponse response, int readTimeout)
        {
            Stream responseStream = response.GetResponseStream();

            responseStream.ReadTimeout = readTimeout;
            return WrapStream(responseStream, response.ContentEncoding);
        }

        private static Stream WrapStream(Stream responseStream, string contentEncoding)
        {
            if (contentEncoding.ToLowerInvariant().Contains("gzip"))
            {
                return new GZipStream(responseStream, CompressionMode.Decompress);
            }
            else if (contentEncoding.ToLowerInvariant().Contains("deflate"))
            {
                return new DeflateStream(responseStream, CompressionMode.Decompress);
            }
            else
            {
                return responseStream;
            }
        }

        #region Methods for subclasses to override

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetXmlElementName();

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetResponseXmlElementName();

        /// <summary>
        /// Gets the minimum server version required to process this request.
        /// </summary>
        /// <returns>Exchange server version.</returns>
        internal abstract ExchangeVersion GetMinimumRequiredServerVersion();

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteElementsToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal abstract object ParseResponse(EwsServiceXmlReader reader);

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="jsonBody">The json body.</param>
        /// <returns>Response object.</returns>
        internal virtual object ParseResponse(JsonObject jsonBody)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            serviceResponse.LoadFromJson(jsonBody, this.Service);
            return serviceResponse;
        }

        /// <summary>
        /// Gets a value indicating whether the TimeZoneContext SOAP header should be eimitted.
        /// </summary>
        /// <value><c>true</c> if the time zone should be emitted; otherwise, <c>false</c>.</value>
        internal virtual bool EmitTimeZoneHeader
        {
            get { return false; }
        }

        #endregion

        /// <summary>
        /// Validate request.
        /// </summary>
        internal virtual void Validate()
        {
            this.Service.Validate();
        }

        /// <summary>
        /// Writes XML body.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteBodyToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Messages, this.GetXmlElementName());

            this.WriteAttributesToXml(writer);
            this.WriteElementsToXml(writer);

            writer.WriteEndElement(); // m:this.GetXmlElementName()
        }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <remarks>
        /// Subclass will override if it has XML attributes.
        /// </remarks>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceRequestBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal ServiceRequestBase(ExchangeService service)
        {
            this.service = service;
            this.ThrowIfNotSupportedByRequestedServerVersion();
        }

        /// <summary>
        /// Gets the service.
        /// </summary>
        /// <value>The service.</value>
        internal ExchangeService Service
        {
            get { return this.service; }
        }

        /// <summary>
        /// Throw exception if request is not supported in requested server version.
        /// </summary>
        /// <exception cref="ServiceVersionException">Raised if request requires a later version of Exchange.</exception>
        internal void ThrowIfNotSupportedByRequestedServerVersion()
        {
            if (this.Service.RequestedServerVersion < this.GetMinimumRequiredServerVersion())
            {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.RequestIncompatibleWithRequestVersion,
                        this.GetXmlElementName(),
                        this.GetMinimumRequiredServerVersion()));
            }
        }

        #region HttpWebRequest-based implementation

        /// <summary>
        /// Writes XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsXmlSchemaInstanceNamespacePrefix, EwsUtilities.EwsXmlSchemaInstanceNamespace);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsMessagesNamespacePrefix, EwsUtilities.EwsMessagesNamespace);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsTypesNamespacePrefix, EwsUtilities.EwsTypesNamespace);
            if (writer.RequireWSSecurityUtilityNamespace)
            {
                writer.WriteAttributeValue("xmlns", EwsUtilities.WSSecurityUtilityNamespacePrefix, EwsUtilities.WSSecurityUtilityNamespace);
            }

            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);

            if (this.Service.Credentials != null)
            {
                this.Service.Credentials.EmitExtraSoapHeaderNamespaceAliases(writer.InternalWriter);
            }

            // Emit the RequestServerVersion header
            if (!this.Service.SuppressXmlVersionHeader)
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.RequestServerVersion);
                writer.WriteAttributeValue(XmlAttributeNames.Version, this.GetRequestedServiceVersionString());
                writer.WriteEndElement(); // RequestServerVersion
            }

            // Against Exchange 2007 SP1, we always emit the simplified time zone header. It adds very little to
            // the request, so bandwidth consumption is not an issue. Against Exchange 2010 and above, we emit
            // the full time zone header but only when the request actually needs it.
            //
            // The exception to this is if we are in Exchange2007 Compat Mode, in which case we should never emit 
            // the header.  (Note: Exchange2007 Compat Mode is enabled for testability purposes only.)
            //
            if ((this.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1 || this.EmitTimeZoneHeader) &&
                (!this.Service.Exchange2007CompatibilityMode))
            {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TimeZoneContext);

                this.Service.TimeZoneDefinition.WriteToXml(writer);

                writer.WriteEndElement(); // TimeZoneContext

                writer.IsTimeZoneHeaderEmitted = true;
            }

            // Emit the MailboxCulture header
            if (this.Service.PreferredCulture != null)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MailboxCulture,
                    this.Service.PreferredCulture.Name);
            }

            // Emit the DateTimePrecision header
            if (this.Service.DateTimePrecision != DateTimePrecision.Default)
            {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DateTimePrecision,
                    this.Service.DateTimePrecision.ToString());
            }

            // Emit the ExchangeImpersonation header
            if (this.Service.ImpersonatedUserId != null)
            {
                this.Service.ImpersonatedUserId.WriteToXml(writer);
            }
            else if (this.Service.PrivilegedUserId != null)
            {
                this.Service.PrivilegedUserId.WriteToXml(writer, this.Service.RequestedServerVersion);
            }
            else if (this.Service.ManagementRoles != null)
            {
                this.Service.ManagementRoles.WriteToXml(writer);
            }

            if (this.Service.Credentials != null)
            {
                this.Service.Credentials.SerializeExtraSoapHeaders(writer.InternalWriter, this.GetXmlElementName());
            }

            this.Service.DoOnSerializeCustomSoapHeaders(writer.InternalWriter);

            writer.WriteEndElement(); // soap:Header

            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);

            this.WriteBodyToXml(writer);

            writer.WriteEndElement(); // soap:Body
            writer.WriteEndElement(); // soap:Envelope
        }

        /// <summary>
        /// Creates the json request.
        /// </summary>
        /// <returns></returns>
        internal JsonObject CreateJsonRequest()
        {
            IJsonSerializable serializableRequest = this as IJsonSerializable;

            if (serializableRequest == null)
            {
                throw new JsonSerializationNotImplementedException();
            }

            JsonObject jsonRequest = new JsonObject();

            jsonRequest.Add("Header", this.CreateJsonHeaders());
            jsonRequest.Add("Body", serializableRequest.ToJson(service));

            return jsonRequest;
        }

        /// <summary>
        /// Creates the json headers.
        /// </summary>
        /// <returns></returns>
        private JsonObject CreateJsonHeaders()
        {
            JsonObject headers = new JsonObject();

            headers.Add(XmlElementNames.RequestServerVersion, this.GetRequestedServiceVersionString());

            // Against Exchange 2007 SP1, we always emit the simplified time zone header. It adds very little to
            // the request, so bandwidth consumption is not an issue. Against Exchange 2010 and above, we emit
            // the full time zone header but only when the request actually needs it.
            //
            // The exception to this is if we are in Exchange2007 Compat Mode, in which case we should never emit 
            // the header.  (Note: Exchange2007 Compat Mode is enabled for testability purposes only.)
            //
            if ((this.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1 || this.EmitTimeZoneHeader) &&
                (!this.Service.Exchange2007CompatibilityMode))
            {
                JsonObject jsonTimeZoneDefinition = new JsonObject();
                jsonTimeZoneDefinition.Add(XmlElementNames.TimeZoneDefinition, this.Service.TimeZoneDefinition.InternalToJson(this.Service));
                headers.Add(XmlElementNames.TimeZoneContext, jsonTimeZoneDefinition);
            }

            if (this.Service.PreferredCulture != null)
            {
                headers.Add(XmlElementNames.MailboxCulture, this.Service.PreferredCulture.Name);
            }

            // Emit the DateTimePrecision header
            if (this.Service.DateTimePrecision != DateTimePrecision.Default)
            {
                headers.Add(XmlElementNames.DateTimePrecision, this.Service.DateTimePrecision.ToString());
            }

            // TODO: JSON-ify the ImpersonatedUserId
            ////// Emit the ExchangeImpersonation header
            ////if (this.Service.ImpersonatedUserId != null)
            ////{
            ////    this.Service.ImpersonatedUserId.WriteToXml(writer);
            ////}

            // TODO: JSON-ify the Credentials
            ////if (this.Service.Credentials != null)
            ////{
            ////    this.Service.Credentials.SerializeExtraSoapHeaders(writer.InternalWriter, this.GetXmlElementName());
            ////}

            if (this.Service.ManagementRoles != null)
            {
                headers.Add(XmlElementNames.ManagementRole, this.Service.ManagementRoles.ToJsonObject());
            }

            return headers;
        }

        /// <summary>
        /// Gets string representation of requested server version.
        /// </summary>
        /// <remarks>
        /// In order to support E12 RTM servers, ExchangeService has another flag indicating that
        /// we should use "Exchange2007" as the server version string rather than Exchange2007_SP1.
        /// </remarks>
        /// <returns>String representation of requested server version.</returns>
        private string GetRequestedServiceVersionString()
        {
            if (this.Service.Exchange2007CompatibilityMode && this.Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
            {
                return "Exchange2007";
            }
            else
            {
                return this.Service.RequestedServerVersion.ToString();
            }
        }

        /// <summary>
        /// Emits the request.
        /// </summary>
        /// <param name="request">The request.</param>
        private void EmitRequest(IEwsHttpWebRequest request)
        {
            if (this.Service.RenderingMethod == ExchangeService.RenderingMode.Xml)
            {
                using (Stream requestStream = this.GetWebRequestStream(request))
                {
                    using (EwsServiceXmlWriter writer = new EwsServiceXmlWriter(this.Service, requestStream))
                    {
                        this.WriteToXml(writer);
                    }
                }
            }
            else if (this.Service.RenderingMethod == ExchangeService.RenderingMode.JSON)
            {
                JsonObject requestObject = this.CreateJsonRequest();

                using (Stream serviceRequestStream = this.GetWebRequestStream(request))
                {
                    requestObject.SerializeToJson(serviceRequestStream);
                }
            }
        }

        /// <summary>
        /// Traces the and emits the request.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="needSignature"></param>
        /// <param name="needTrace"></param>
        private void TraceAndEmitRequest(IEwsHttpWebRequest request, bool needSignature, bool needTrace)
        {
            if (this.service.RenderingMethod == ExchangeService.RenderingMode.Xml)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (EwsServiceXmlWriter writer = new EwsServiceXmlWriter(this.Service, memoryStream))
                    {
                        writer.RequireWSSecurityUtilityNamespace = needSignature;
                        this.WriteToXml(writer);
                    }

                    if (needSignature)
                    {
                        this.service.Credentials.Sign(memoryStream);
                    }

                    if (needTrace)
                    {
                        this.TraceXmlRequest(memoryStream);
                    }

                    using (Stream serviceRequestStream = this.GetWebRequestStream(request))
                    {
                        EwsUtilities.CopyStream(memoryStream, serviceRequestStream);
                    }
                }
            }
            else if (this.service.RenderingMethod == ExchangeService.RenderingMode.JSON)
            {
                JsonObject requestObject = this.CreateJsonRequest();

                this.TraceJsonRequest(requestObject);

                using (Stream serviceRequestStream = this.GetWebRequestStream(request))
                {
                    requestObject.SerializeToJson(serviceRequestStream);
                }
            }
        }

        /// <summary>
        /// Get the request stream
        /// </summary>
        /// <param name="request">The request</param>
        /// <returns>The Request stream</returns>
        private Stream GetWebRequestStream(IEwsHttpWebRequest request)
        {
            // In the async case, although we can use async callback to make the entire worflow completely async, 
            // there is little perf gain with this approach because of EWS's message nature.
            // The overall latency of BeginGetRequestStream() is same as GetRequestStream() in this case.
            // The overhead to implement a two-step async operation includes wait handle synchronization, exception handling and wrapping.
            // Therefore, we only leverage BeginGetResponse() and EndGetReponse() to provide the async functionality.
            // Reference: http://www.wintellect.com/CS/blogs/jeffreyr/archive/2009/02/08/httpwebrequest-its-request-stream-and-sending-data-in-chunks.aspx
            return request.EndGetRequestStream(request.BeginGetRequestStream(null, null));
        }

        /// <summary>
        /// Reads the response.
        /// </summary>
        /// <param name="ewsXmlReader">The XML reader.</param>
        /// <returns>Service response.</returns>
        protected object ReadResponse(EwsServiceXmlReader ewsXmlReader)
        {
            object serviceResponse;

            this.ReadPreamble(ewsXmlReader);
            ewsXmlReader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            this.ReadSoapHeader(ewsXmlReader);
            ewsXmlReader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);

            ewsXmlReader.ReadStartElement(XmlNamespace.Messages, this.GetResponseXmlElementName());

            serviceResponse = this.ParseResponse(ewsXmlReader);

            ewsXmlReader.ReadEndElementIfNecessary(XmlNamespace.Messages, this.GetResponseXmlElementName());

            ewsXmlReader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);
            ewsXmlReader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            return serviceResponse;
        }

        /// <summary>
        /// Builds the response object from json.
        /// </summary>
        /// <param name="jsonResponse">The json response.</param>
        /// <returns></returns>
        protected object BuildResponseObjectFromJson(JsonObject jsonResponse)
        {
            if (jsonResponse.ContainsKey("Header"))
            {
                this.ReadSoapHeader(jsonResponse.ReadAsJsonObject("Header"));
            }

            return this.ParseResponse(jsonResponse.ReadAsJsonObject(XmlElementNames.SOAPBodyElementName));
        }

        /// <summary>
        /// Reads any preamble data not part of the core response.
        /// </summary>
        /// <param name="ewsXmlReader">The EwsServiceXmlReader.</param>
        protected virtual void ReadPreamble(EwsServiceXmlReader ewsXmlReader)
        {
            this.ReadXmlDeclaration(ewsXmlReader);
        }

        /// <summary>
        /// Read SOAP header and extract server version
        /// </summary>
        /// <param name="reader">EwsServiceXmlReader</param>
        private void ReadSoapHeader(EwsServiceXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);
            do
            {
                reader.Read();

                // Is this the ServerVersionInfo?
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ServerVersionInfo))
                {
                    this.Service.ServerInfo = ExchangeServerInfo.Parse(reader);
                }

                // Ignore anything else inside the SOAP header
            }
            while (!reader.IsEndElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName));
        }

        /// <summary>
        /// Read SOAP header and extract server version
        /// </summary>
        /// <param name="jsonHeader">The json header.</param>
        private void ReadSoapHeader(JsonObject jsonHeader)
        {
            if (jsonHeader.ContainsKey(XmlElementNames.ServerVersionInfo))
            {
                this.Service.ServerInfo = ExchangeServerInfo.Parse(jsonHeader.ReadAsJsonObject(XmlElementNames.ServerVersionInfo));
            }
        }

        /// <summary>
        /// Reads the SOAP fault.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>SOAP fault details.</returns>
        protected SoapFaultDetails ReadSoapFault(EwsServiceXmlReader reader)
        {
            SoapFaultDetails soapFaultDetails = null;

            try
            {
                this.ReadXmlDeclaration(reader);

                reader.Read();
                if (!reader.IsStartElement() || (reader.LocalName != XmlElementNames.SOAPEnvelopeElementName))
                {
                    return soapFaultDetails;
                }

                // EWS can sometimes return SOAP faults using the SOAP 1.2 namespace. Get the
                // namespace URI from the envelope element and use it for the rest of the parsing.
                // If it's not 1.1 or 1.2, we can't continue.
                XmlNamespace soapNamespace = EwsUtilities.GetNamespaceFromUri(reader.NamespaceUri);
                if (soapNamespace == XmlNamespace.NotSpecified)
                {
                    return soapFaultDetails;
                }

                reader.Read();

                // EWS doesn't always return a SOAP header. If this response contains a header element, 
                // read the server version information contained in the header.
                if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPHeaderElementName))
                {
                    do
                    {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ServerVersionInfo))
                        {
                            this.Service.ServerInfo = ExchangeServerInfo.Parse(reader);
                        }
                    }
                    while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPHeaderElementName));

                    // Queue up the next read
                    reader.Read();
                }

                // Parse the fault element contained within the SOAP body.
                if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPBodyElementName))
                {
                    do
                    {
                        reader.Read();

                        // Parse Fault element
                        if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPFaultElementName))
                        {
                            soapFaultDetails = SoapFaultDetails.Parse(reader, soapNamespace);
                        }
                    }
                    while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPBodyElementName));
                }

                reader.ReadEndElement(soapNamespace, XmlElementNames.SOAPEnvelopeElementName);
            }
            catch (XmlException)
            {
                // If response doesn't contain a valid SOAP fault, just ignore exception and
                // return null for SOAP fault details.
            }

            return soapFaultDetails;
        }

        /// <summary>
        /// Reads the SOAP fault.
        /// </summary>
        /// <param name="jsonSoapFault">The json SOAP fault.</param>
        /// <returns></returns>
        private SoapFaultDetails ReadSoapFault(JsonObject jsonSoapFault)
        {
            SoapFaultDetails soapFaultDetails = null;

            if (jsonSoapFault.ContainsKey("Header"))
            {
                this.ReadSoapHeader(jsonSoapFault.ReadAsJsonObject("Header"));
            }

            if (jsonSoapFault.ContainsKey("Body"))
            {
                soapFaultDetails = SoapFaultDetails.Parse(jsonSoapFault.ReadAsJsonObject("Body"));
            }

            return soapFaultDetails;
        }

        /// <summary>
        /// Validates request parameters, and emits the request to the server.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns>The response returned by the server.</returns>
        protected IEwsHttpWebResponse ValidateAndEmitRequest(out IEwsHttpWebRequest request)
        {
            this.Validate();

            request = this.BuildEwsHttpWebRequest();

            if (this.service.SendClientLatencies)
            {
                string clientStatisticsToAdd = null;

                lock (clientStatisticsCache)
                {
                    if (clientStatisticsCache.Count > 0)
                    {
                        clientStatisticsToAdd = clientStatisticsCache[0];
                        clientStatisticsCache.RemoveAt(0);
                    }
                }

                if (!string.IsNullOrEmpty(clientStatisticsToAdd))
                {
                    if (request.Headers[ClientStatisticsRequestHeader] != null)
                    {
                        request.Headers[ClientStatisticsRequestHeader] =
                            request.Headers[ClientStatisticsRequestHeader]
                            + clientStatisticsToAdd;
                    }
                    else
                    {
                        request.Headers.Add(
                            ClientStatisticsRequestHeader,
                            clientStatisticsToAdd);
                    }
                }
            }

            DateTime startTime = DateTime.UtcNow;
            IEwsHttpWebResponse response = null;

            try
            {
                response = this.GetEwsHttpWebResponse(request);
            }
            finally
            {
                if (this.service.SendClientLatencies)
                {
                    int clientSideLatency = (int)(DateTime.UtcNow - startTime).TotalMilliseconds;
                    string requestId = string.Empty;
                    string soapAction = this.GetType().Name.Replace("Request", string.Empty);

                    if (response != null && response.Headers != null)
                    {
                        foreach (string requestIdHeader in ServiceRequestBase.RequestIdResponseHeaders)
                        {
                            string requestIdValue = response.Headers.Get(requestIdHeader);
                            if (!string.IsNullOrEmpty(requestIdValue))
                            {
                                requestId = requestIdValue;
                                break;
                            }
                        }
                    }

                    StringBuilder sb = new StringBuilder();
                    sb.Append("MessageId=");
                    sb.Append(requestId);
                    sb.Append(",ResponseTime=");
                    sb.Append(clientSideLatency);
                    sb.Append(",SoapAction=");
                    sb.Append(soapAction);
                    sb.Append(";");

                    lock (clientStatisticsCache)
                    {
                        clientStatisticsCache.Add(sb.ToString());
                    }
                }
            }

            return response;
        }

        /// <summary>
        /// Builds the IEwsHttpWebRequest object for current service request with exception handling.
        /// </summary>
        /// <returns>An IEwsHttpWebRequest instance</returns>
        protected IEwsHttpWebRequest BuildEwsHttpWebRequest()
        {
            try
            {
                IEwsHttpWebRequest request = this.Service.PrepareHttpWebRequest(this.GetXmlElementName());

                this.Service.TraceHttpRequestHeaders(TraceFlags.EwsRequestHttpHeaders, request);

                bool needSignature = this.Service.Credentials != null && this.Service.Credentials.NeedSignature;
                bool needTrace = this.Service.IsTraceEnabledFor(TraceFlags.EwsRequest);

                // If tracing is enabled, we generate the request in-memory so that we
                // can pass it along to the ITraceListener. Then we copy the stream to
                // the request stream.
                if (needSignature || needTrace)
                {
                    this.TraceAndEmitRequest(request, needSignature, needTrace);
                }
                else
                {
                    this.EmitRequest(request);
                }

                return request;
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    this.ProcessWebException(ex);
                }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
            }
            catch (IOException e)
            {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
        }

        /// <summary>
        ///  Gets the IEwsHttpWebRequest object from the specified IEwsHttpWebRequest object with exception handling
        /// </summary>
        /// <param name="request">The specified IEwsHttpWebRequest</param>
        /// <returns>An IEwsHttpWebResponse instance</returns>
        protected IEwsHttpWebResponse GetEwsHttpWebResponse(IEwsHttpWebRequest request)
        {
            try
            {
                return request.GetResponse();
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    this.ProcessWebException(ex);
                }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
            }
            catch (IOException e)
            {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
        }

        /// <summary>
        /// Ends getting the specified async IEwsHttpWebRequest object from the specified IEwsHttpWebRequest object with exception handling.
        /// </summary>
        /// <param name="request">The specified IEwsHttpWebRequest</param>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>An IEwsHttpWebResponse instance</returns>
        protected IEwsHttpWebResponse EndGetEwsHttpWebResponse(IEwsHttpWebRequest request, IAsyncResult asyncResult)
        {
            try
            {
                // Note that this call may throw ArgumentException if the HttpWebRequest instance is not the original one,
                // and we just let it out
                return request.EndGetResponse(asyncResult);
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    this.ProcessWebException(ex);
                }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
            }
            catch (IOException e)
            {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
        }

        /// <summary>
        /// Processes the web exception.
        /// </summary>
        /// <param name="webException">The web exception.</param>
        private void ProcessWebException(WebException webException)
        {
            if (webException.Response != null)
            {
                IEwsHttpWebResponse httpWebResponse = this.Service.HttpWebRequestFactory.CreateExceptionResponse(webException);
                SoapFaultDetails soapFaultDetails = null;

                if (httpWebResponse.StatusCode == HttpStatusCode.InternalServerError)
                {
                    this.Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, httpWebResponse);

                    // If tracing is enabled, we read the entire response into a MemoryStream so that we
                    // can pass it along to the ITraceListener. Then we parse the response from the 
                    // MemoryStream.
                    if (this.Service.IsTraceEnabledFor(TraceFlags.EwsResponse))
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            using (Stream serviceResponseStream = ServiceRequestBase.GetResponseStream(httpWebResponse))
                            {
                                // Copy response to in-memory stream and reset position to start.
                                EwsUtilities.CopyStream(serviceResponseStream, memoryStream);
                                memoryStream.Position = 0;
                            }

                            if (this.Service.RenderingMethod == ExchangeService.RenderingMode.Xml)
                            {
                                this.TraceResponseXml(httpWebResponse, memoryStream);

                                EwsServiceXmlReader reader = new EwsServiceXmlReader(memoryStream, this.Service);
                                soapFaultDetails = this.ReadSoapFault(reader);
                            }
                            else if (this.Service.RenderingMethod == ExchangeService.RenderingMode.JSON)
                            {
                                this.TraceResponseJson(httpWebResponse, memoryStream);

                                try
                                {
                                    JsonObject jsonSoapFault = new JsonParser(memoryStream).Parse();
                                    soapFaultDetails = this.ReadSoapFault(jsonSoapFault);
                                }
                                catch (ServiceJsonDeserializationException)
                                {
                                    // If no valid JSON response was returned, just return null SoapFault details
                                }
                            }
                            else
                            {
                                throw new InvalidOperationException();
                            }
                        }
                    }
                    else
                    {
                        using (Stream stream = ServiceRequestBase.GetResponseStream(httpWebResponse))
                        {
                            if (this.Service.RenderingMethod == ExchangeService.RenderingMode.Xml)
                            {
                                EwsServiceXmlReader reader = new EwsServiceXmlReader(stream, this.Service);
                                soapFaultDetails = this.ReadSoapFault(reader);
                            }
                            else if (this.Service.RenderingMethod == ExchangeService.RenderingMode.JSON)
                            {
                                try
                                {
                                    JsonObject jsonSoapFault = new JsonParser(stream).Parse();
                                    soapFaultDetails = this.ReadSoapFault(jsonSoapFault);
                                }
                                catch (ServiceJsonDeserializationException)
                                {
                                    // If no valid JSON response was returned, just return null SoapFault details
                                }
                            }
                            else
                            {
                                throw new InvalidOperationException();
                            }
                        }
                    }

                    if (soapFaultDetails != null)
                    {
                        switch (soapFaultDetails.ResponseCode)
                        {
                            case ServiceError.ErrorInvalidServerVersion:
                                throw new ServiceVersionException(Strings.ServerVersionNotSupported);

                            case ServiceError.ErrorSchemaValidation:
                                // If we're talking to an E12 server (8.00.xxxx.xxx), a schema validation error is the same as a version mismatch error.
                                // (Which only will happen if we send a request that's not valid for E12).
                                if ((this.Service.ServerInfo != null) &&
                                    (this.Service.ServerInfo.MajorVersion == 8) && (this.Service.ServerInfo.MinorVersion == 0))
                                {
                                    throw new ServiceVersionException(Strings.ServerVersionNotSupported);
                                }

                                break;

                            case ServiceError.ErrorIncorrectSchemaVersion:
                                // This shouldn't happen. It indicates that a request wasn't valid for the version that was specified.
                                EwsUtilities.Assert(
                                    false,
                                    "ServiceRequestBase.ProcessWebException",
                                    "Exchange server supports requested version but request was invalid for that version");
                                break;

                            case ServiceError.ErrorServerBusy:
                                throw new ServerBusyException(new ServiceResponse(soapFaultDetails));

                            default:
                                // Other error codes will be reported as remote error
                                break;
                        }

                        // General fall-through case: throw a ServiceResponseException
                        throw new ServiceResponseException(new ServiceResponse(soapFaultDetails));
                    }
                }
                else
                {
                    this.Service.ProcessHttpErrorResponse(httpWebResponse, webException);
                }
            }
        }

        /// <summary>
        /// Traces an XML request.  This should only be used for synchronous requests, or synchronous situations
        /// (such as a WebException on an asynchrounous request).
        /// </summary>
        /// <param name="memoryStream">The request content in a MemoryStream.</param>
        protected void TraceXmlRequest(MemoryStream memoryStream)
        {
            this.Service.TraceXml(TraceFlags.EwsRequest, memoryStream);
        }

        /// <summary>
        /// Traces a JSON request. This should only be used for synchronous requests, or synchronous situations
        /// (such as a WebException on an asynchrounous request).
        /// </summary>
        /// <param name="requestObject">The JSON request object.</param>
        protected void TraceJsonRequest(JsonObject requestObject)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                requestObject.SerializeToJson(memoryStream, this.Service.TraceEnablePrettyPrinting);

                memoryStream.Position = 0;

                using (StreamReader reader = new StreamReader(memoryStream))
                {
                    this.Service.TraceMessage(TraceFlags.EwsRequest, reader.ReadToEnd());
                }
            }
        }

        /// <summary>
        /// Traces the response.  This should only be used for synchronous requests, or synchronous situations
        /// (such as a WebException on an asynchrounous request).
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="memoryStream">The response content in a MemoryStream.</param>
        protected void TraceResponseXml(IEwsHttpWebResponse response, MemoryStream memoryStream)
        {
            if (!string.IsNullOrEmpty(response.ContentType) &&
                (response.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) ||
                 response.ContentType.StartsWith("application/soap", StringComparison.OrdinalIgnoreCase)))
            {
                this.Service.TraceXml(TraceFlags.EwsResponse, memoryStream);
            }
            else
            {
                this.Service.TraceMessage(TraceFlags.EwsResponse, "Non-textual response");
            }
        }

        /// <summary>
        /// Traces the response.  This should only be used for synchronous requests, or synchronous situations
        /// (such as a WebException on an asynchrounous request).
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="memoryStream">The response content in a MemoryStream.</param>
        protected void TraceResponseJson(IEwsHttpWebResponse response, MemoryStream memoryStream)
        {
            JsonObject jsonResponse = new JsonParser(memoryStream).Parse();

            using (MemoryStream responseStream = new MemoryStream())
            {
                jsonResponse.SerializeToJson(responseStream, this.Service.TraceEnablePrettyPrinting);

                responseStream.Position = 0;

                using (StreamReader responseReader = new StreamReader(responseStream))
                {
                    this.Service.TraceMessage(TraceFlags.EwsResponse, responseReader.ReadToEnd());
                }
            }

            memoryStream.Seek(0, SeekOrigin.Begin);
        }

        /// <summary>
        /// Try to read the XML declaration. If it's not there, the server didn't return XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadXmlDeclaration(EwsServiceXmlReader reader)
        {
            try
            {
                reader.Read(XmlNodeType.XmlDeclaration);
            }
            catch (XmlException ex)
            {
                throw new ServiceRequestException(Strings.ServiceResponseDoesNotContainXml, ex);
            }
            catch (ServiceXmlDeserializationException ex)
            {
                throw new ServiceRequestException(Strings.ServiceResponseDoesNotContainXml, ex);
            }
        }

        #endregion
    }
}
