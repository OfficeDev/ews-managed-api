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
// <summary>Defines the AutodiscoverRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.IO;
    using System.IO.Compression;
    using System.Net;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the base class for all requested made to the Autodiscover service.
    /// </summary>
    internal abstract class AutodiscoverRequest
    {
        private AutodiscoverService service;
        private Uri url;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverRequest"/> class.
        /// </summary>
        /// <param name="service">Autodiscover service associated with this request.</param>
        /// <param name="url">URL of Autodiscover service.</param>
        internal AutodiscoverRequest(AutodiscoverService service, Uri url)
        {
            this.service = service;
            this.url = url;
        }

        /// <summary>
        /// Determines whether response is a redirection.
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        /// <returns>True if redirection response.</returns>
        internal static bool IsRedirectionResponse(IEwsHttpWebResponse httpWebResponse)
        {
            return (httpWebResponse.StatusCode == HttpStatusCode.Redirect) ||
                   (httpWebResponse.StatusCode == HttpStatusCode.Moved) ||
                   (httpWebResponse.StatusCode == HttpStatusCode.RedirectKeepVerb) ||
                   (httpWebResponse.StatusCode == HttpStatusCode.RedirectMethod);
        }

        /// <summary>
        /// Validates the request.
        /// </summary>
        internal virtual void Validate()
        {
            this.Service.Validate();
        }

        /// <summary>
        /// Executes this instance.
        /// </summary>
        /// <returns></returns>
        internal AutodiscoverResponse InternalExecute()
        {
            this.Validate();

            try
            {
                IEwsHttpWebRequest request = this.Service.PrepareHttpWebRequestForUrl(this.Url);

                this.Service.TraceHttpRequestHeaders(TraceFlags.AutodiscoverRequestHttpHeaders, request);

                bool needSignature = this.Service.Credentials != null && this.Service.Credentials.NeedSignature;
                bool needTrace = this.Service.IsTraceEnabledFor(TraceFlags.AutodiscoverRequest);

                using (Stream requestStream = request.GetRequestStream())
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        using (EwsServiceXmlWriter writer = new EwsServiceXmlWriter(this.Service, memoryStream))
                        {
                            writer.RequireWSSecurityUtilityNamespace = needSignature;
                            this.WriteSoapRequest(
                                this.Url, 
                                writer);
                        }

                        if (needSignature)
                        {
                            this.service.Credentials.Sign(memoryStream);
                        }

                        if (needTrace)
                        {
                            memoryStream.Position = 0;
                            this.Service.TraceXml(TraceFlags.AutodiscoverRequest, memoryStream);
                        }

                        EwsUtilities.CopyStream(memoryStream, requestStream);
                    }
                }

                using (IEwsHttpWebResponse webResponse = request.GetResponse())
                {
                    if (AutodiscoverRequest.IsRedirectionResponse(webResponse))
                    {
                        AutodiscoverResponse response = this.CreateRedirectionResponse(webResponse);
                        if (response != null)
                        {
                            return response;
                        }
                        else
                        {
                            throw new ServiceRemoteException(Strings.InvalidRedirectionResponseReturned);
                        }
                    }

                    using (Stream responseStream = AutodiscoverRequest.GetResponseStream(webResponse))
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            // Copy response stream to in-memory stream and reset to start
                            EwsUtilities.CopyStream(responseStream, memoryStream);
                            memoryStream.Position = 0;

                            this.Service.TraceResponse(webResponse, memoryStream);                           

                            EwsXmlReader ewsXmlReader = new EwsXmlReader(memoryStream);

                            // WCF may not generate an XML declaration.
                            ewsXmlReader.Read();
                            if (ewsXmlReader.NodeType == XmlNodeType.XmlDeclaration)
                            {
                                ewsXmlReader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
                            }
                            else if ((ewsXmlReader.NodeType != XmlNodeType.Element) || (ewsXmlReader.LocalName != XmlElementNames.SOAPEnvelopeElementName) || (ewsXmlReader.NamespaceUri != EwsUtilities.GetNamespaceUri(XmlNamespace.Soap)))
                            {
                                throw new ServiceXmlDeserializationException(Strings.InvalidAutodiscoverServiceResponse);
                            }

                            this.ReadSoapHeaders(ewsXmlReader);

                            AutodiscoverResponse response = this.ReadSoapBody(ewsXmlReader);

                            ewsXmlReader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);

                            if (response.ErrorCode == AutodiscoverErrorCode.NoError)
                            {
                                return response;
                            }
                            else
                            {
                                throw new AutodiscoverResponseException(response.ErrorCode, response.ErrorMessage);
                            }
                        }
                    }
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                {
                    IEwsHttpWebResponse httpWebResponse = this.Service.HttpWebRequestFactory.CreateExceptionResponse(ex);

                    if (AutodiscoverRequest.IsRedirectionResponse(httpWebResponse))
                    {
                        this.Service.ProcessHttpResponseHeaders(
                            TraceFlags.AutodiscoverResponseHttpHeaders,
                            httpWebResponse);

                        AutodiscoverResponse response = this.CreateRedirectionResponse(httpWebResponse);
                        if (response != null)
                        {
                            return response;
                        }
                    }
                    else
                    {
                        this.ProcessWebException(ex);
                    }
                }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
            }
            catch (XmlException ex)
            {
                this.Service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("XML parsing error: {0}", ex.Message));

                // Wrap exception
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
            }
            catch (IOException ex)
            {
                this.Service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("I/O error: {0}", ex.Message));

                // Wrap exception
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
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
                SoapFaultDetails soapFaultDetails;

                if (httpWebResponse.StatusCode == HttpStatusCode.InternalServerError)
                {
                    // If tracing is enabled, we read the entire response into a MemoryStream so that we
                    // can pass it along to the ITraceListener. Then we parse the response from the 
                    // MemoryStream.
                    if (this.Service.IsTraceEnabledFor(TraceFlags.AutodiscoverRequest))
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            using (Stream serviceResponseStream = AutodiscoverRequest.GetResponseStream(httpWebResponse))
                            {
                                // Copy response to in-memory stream and reset position to start.
                                EwsUtilities.CopyStream(serviceResponseStream, memoryStream);
                                memoryStream.Position = 0;
                            }

                            this.Service.TraceResponse(httpWebResponse, memoryStream);

                            EwsXmlReader reader = new EwsXmlReader(memoryStream);
                            soapFaultDetails = this.ReadSoapFault(reader);
                        }
                    }
                    else
                    {
                        using (Stream stream = AutodiscoverRequest.GetResponseStream(httpWebResponse))
                        {
                            EwsXmlReader reader = new EwsXmlReader(stream);
                            soapFaultDetails = this.ReadSoapFault(reader);
                        }
                    }

                    if (soapFaultDetails != null)
                    {
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
        /// Create a redirection response.
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        private AutodiscoverResponse CreateRedirectionResponse(IEwsHttpWebResponse httpWebResponse)
        {
            string location = httpWebResponse.Headers[HttpResponseHeader.Location];
            if (!string.IsNullOrEmpty(location))
            {
                try
                {
                    Uri redirectionUri = new Uri(this.Url, location);
                    if ((redirectionUri.Scheme == Uri.UriSchemeHttp) || (redirectionUri.Scheme == Uri.UriSchemeHttps))
                    {
                        AutodiscoverResponse response = this.CreateServiceResponse();
                        response.ErrorCode = AutodiscoverErrorCode.RedirectUrl;
                        response.RedirectionUrl = redirectionUri;
                        return response;
                    }

                    this.Service.TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("Invalid redirection URL '{0}' returned by Autodiscover service.", redirectionUri));
                }
                catch (UriFormatException)
                {
                    this.Service.TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("Invalid redirection location '{0}' returned by Autodiscover service.", location));
                }
            }
            else
            {
                this.Service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    "Redirection response returned by Autodiscover service without redirection location.");
            }

            return null;
        }

        /// <summary>
        /// Reads the SOAP fault.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>SOAP fault details.</returns>
        private SoapFaultDetails ReadSoapFault(EwsXmlReader reader)
        {
            SoapFaultDetails soapFaultDetails = null;

            try
            {
                // WCF may not generate an XML declaration.
                reader.Read();
                if (reader.NodeType == XmlNodeType.XmlDeclaration)
                {
                    reader.Read();
                }

                if (!reader.IsStartElement() || (reader.LocalName != XmlElementNames.SOAPEnvelopeElementName))
                {
                    return soapFaultDetails;
                }

                // Get the namespace URI from the envelope element and use it for the rest of the parsing.
                // If it's not 1.1 or 1.2, we can't continue.
                XmlNamespace soapNamespace = EwsUtilities.GetNamespaceFromUri(reader.NamespaceUri);
                if (soapNamespace == XmlNamespace.NotSpecified)
                {
                    return soapFaultDetails;
                }

                reader.Read();

                // Skip SOAP header.
                if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPHeaderElementName))
                {
                    do
                    {
                        reader.Read();
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
        /// Writes the autodiscover SOAP request.
        /// </summary>
        /// <param name="requestUrl">Request URL.</param>
        /// <param name="writer">The writer.</param>
        internal void WriteSoapRequest(
            Uri requestUrl,
            EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            writer.WriteAttributeValue("xmlns", EwsUtilities.AutodiscoverSoapNamespacePrefix, EwsUtilities.AutodiscoverSoapNamespace);
            writer.WriteAttributeValue("xmlns", EwsUtilities.WSAddressingNamespacePrefix, EwsUtilities.WSAddressingNamespace);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsXmlSchemaInstanceNamespacePrefix, EwsUtilities.EwsXmlSchemaInstanceNamespace);
            if (writer.RequireWSSecurityUtilityNamespace)
            {
                writer.WriteAttributeValue("xmlns", EwsUtilities.WSSecurityUtilityNamespacePrefix, EwsUtilities.WSSecurityUtilityNamespace);
            }

            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);

            if (this.Service.Credentials != null)
            {
                this.Service.Credentials.EmitExtraSoapHeaderNamespaceAliases(writer.InternalWriter);
            }

            writer.WriteElementValue(
                XmlNamespace.Autodiscover,
                XmlElementNames.RequestedServerVersion,
                this.Service.RequestedServerVersion.ToString());

            writer.WriteElementValue(
                XmlNamespace.WSAddressing,
                XmlElementNames.Action,
               this.GetWsAddressingActionName());

            writer.WriteElementValue(
                XmlNamespace.WSAddressing,
                XmlElementNames.To,
               requestUrl.AbsoluteUri);

            this.WriteExtraCustomSoapHeadersToXml(writer);

            if (this.Service.Credentials != null)
            {
                this.Service.Credentials.SerializeWSSecurityHeaders(writer.InternalWriter);
            }

            this.Service.DoOnSerializeCustomSoapHeaders(writer.InternalWriter);

            writer.WriteEndElement(); // soap:Header

            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);

            this.WriteBodyToXml(writer);

            writer.WriteEndElement(); // soap:Body
            writer.WriteEndElement(); // soap:Envelope
            writer.Flush();
        }

        /// <summary>
        /// Write extra headers. 
        /// </summary>
        /// <param name="writer">The writer</param>
        internal virtual void WriteExtraCustomSoapHeadersToXml(EwsServiceXmlWriter writer)
        {
            // do nothing here. 
            // currently used only by GetUserSettingRequest to emit the BinarySecret header.
        }

        /// <summary>
        /// Writes XML body.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteBodyToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Autodiscover, this.GetRequestXmlElementName());
            this.WriteAttributesToXml(writer);
            this.WriteElementsToXml(writer);

            writer.WriteEndElement(); // m:this.GetXmlElementName()
        }

        /// <summary>
        /// Gets the response stream (may be wrapped with GZip/Deflate stream to decompress content)
        /// </summary>
        /// <param name="response">HttpWebResponse.</param>
        /// <returns>ResponseStream</returns>
        protected static Stream GetResponseStream(IEwsHttpWebResponse response)
        {
            string contentEncoding = response.ContentEncoding;
            Stream responseStream = response.GetResponseStream();

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

        /// <summary>
        /// Read SOAP headers.
        /// </summary>
        /// <param name="reader">EwsXmlReader</param>
        internal void ReadSoapHeaders(EwsXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);
            do
            {
                reader.Read();

                this.ReadSoapHeader(reader);
            }
            while (!reader.IsEndElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName));
        }

        /// <summary>
        /// Reads a single SOAP header.
        /// </summary>
        /// <param name="reader">EwsXmlReader</param>
        internal virtual void ReadSoapHeader(EwsXmlReader reader)
        {
            // Is this the ServerVersionInfo?
            if (reader.IsStartElement(XmlNamespace.Autodiscover, XmlElementNames.ServerVersionInfo))
            {
                this.service.ServerInfo = this.ReadServerVersionInfo(reader);
            }
        }

        /// <summary>
        /// Read ServerVersionInfo SOAP header.
        /// </summary>
        /// <param name="reader">EwsXmlReader</param>
        private ExchangeServerInfo ReadServerVersionInfo(EwsXmlReader reader)
        {
            ExchangeServerInfo serverInfo = new ExchangeServerInfo();
            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.MajorVersion:
                            serverInfo.MajorVersion = reader.ReadElementValue<int>();
                            break;
                        case XmlElementNames.MinorVersion:
                            serverInfo.MinorVersion = reader.ReadElementValue<int>();
                            break;
                        case XmlElementNames.MajorBuildNumber:
                            serverInfo.MajorBuildNumber = reader.ReadElementValue<int>();
                            break;
                        case XmlElementNames.MinorBuildNumber:
                            serverInfo.MinorBuildNumber = reader.ReadElementValue<int>();
                            break;
                        case XmlElementNames.Version:
                            serverInfo.VersionString = reader.ReadElementValue();
                            break;
                        default:
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.ServerVersionInfo));

            return serverInfo;
        }

        /// <summary>
        /// Read SOAP body.
        /// </summary>
        /// <param name="reader">EwsXmlReader</param>
        internal AutodiscoverResponse ReadSoapBody(EwsXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);
            AutodiscoverResponse responses = this.LoadFromXml(reader);
            reader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);
            return responses;
        }

        /// <summary>
        /// Loads responses from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns></returns>
        internal AutodiscoverResponse LoadFromXml(EwsXmlReader reader)
        {
            string elementName = this.GetResponseXmlElementName();
            reader.ReadStartElement(XmlNamespace.Autodiscover, elementName);
            AutodiscoverResponse response = this.CreateServiceResponse();
            response.LoadFromXml(reader, elementName);
            return response;
        }

        /// <summary>
        /// Gets the name of the request XML element.
        /// </summary>
        /// <returns></returns>
        internal abstract string GetRequestXmlElementName();

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns></returns>
        internal abstract string GetResponseXmlElementName();

        /// <summary>
        /// Gets the WS-Addressing action name.
        /// </summary>
        /// <returns></returns>
        internal abstract string GetWsAddressingActionName();

        /// <summary>
        /// Creates the service response.
        /// </summary>
        /// <returns>AutodiscoverResponse</returns>
        internal abstract AutodiscoverResponse CreateServiceResponse();

        /// <summary>
        /// Writes attributes to request XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteAttributesToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Writes elements to request XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteElementsToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Gets the service.
        /// </summary>
        internal AutodiscoverService Service
        {
            get { return this.service; }
        }

        /// <summary>
        /// Gets the URL.
        /// </summary>
        internal Uri Url
        {
            get { return this.url; }
        }
    }
}
