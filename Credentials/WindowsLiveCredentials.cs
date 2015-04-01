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
    using System.Net;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// WindowsLiveCredentials provides credentials for Windows Live ID authentication.
    /// </summary>
    internal sealed class WindowsLiveCredentials : WSSecurityBasedCredentials
    {
        private string windowsLiveId;
        private string password;
        private Uri windowsLiveUrl;
        private bool isAuthenticated;
        private bool traceEnabled;
        private ITraceListener traceListener = new EwsTraceListener();

        // XML-Encryption Namespace.
        internal const string XmlEncNamespace = "http://www.w3.org/2001/04/xmlenc#";

        // Windows Live SOAP namespace prefix (which is S: instead of soap:)
        internal const string WindowsLiveSoapNamespacePrefix = "S";
        
        // XML element names used in RSTR responses from Windows Live
        internal const string RequestSecurityTokenResponseCollectionElementName = "RequestSecurityTokenResponseCollection";
        internal const string RequestSecurityTokenResponseElementName = "RequestSecurityTokenResponse";
        internal const string EncryptedDataElementName = "EncryptedData";
        internal const string PpElementName = "pp";
        internal const string ReqstatusElementName = "reqstatus";

        // The reqstatus we should receive from Windows Live.
        internal const string SuccessfulReqstatus = "0x0";
        
        // The default Windows Live URL.
        internal static readonly Uri DefaultWindowsLiveUrl = new Uri("https://login.live.com/rst2.srf");

        // The reference we use for creating the XML signature.
        internal const string XmlSignatureReference = "_EWSTKREF";

        /// <summary>
        /// Initializes a new instance of the <see cref="WindowsLiveCredentials"/> class.
        /// </summary>
        /// <param name="windowsLiveId">The user's WindowsLiveId.</param>
        /// <param name="password">The password.</param>
        public WindowsLiveCredentials(string windowsLiveId, string password)
        {
            if (windowsLiveId == null)
            {
                throw new ArgumentNullException("windowsLiveId");
            }
            if (password == null)
            {
                throw new ArgumentNullException("password");
            }
            
            this.windowsLiveId = windowsLiveId;
            this.password = password;
            this.windowsLiveUrl = WindowsLiveCredentials.DefaultWindowsLiveUrl;
        }

        /// <summary>
        /// Gets or sets a flag indicating whether tracing is enabled.
        /// </summary>
        public bool TraceEnabled
        {
            get
            { 
                return this.traceEnabled;
            }

            set 
            {
                this.traceEnabled = value;
                if (this.traceEnabled && (this.traceListener == null))
                {
                    this.traceListener = new EwsTraceListener();
                }
            }
        }

        /// <summary>
        /// Gets or sets the trace listener.
        /// </summary>
        /// <value>The trace listener.</value>
        public ITraceListener TraceListener
        {
            get
            {
                return this.traceListener;
            }

            set
            {
                this.traceListener = value;
                this.traceEnabled = value != null;
            }
        }

        /// <summary>
        /// Gets or sets the Windows Live Url to use.
        /// </summary>
        public Uri WindowsLiveUrl
        {
            get
            {
                return this.windowsLiveUrl;
            }

            set
            {
                // Reset the EWS URL to make sure we go back and re-authenticate next time.
                this.EwsUrl = null;
                this.IsAuthenticated = false;
                this.windowsLiveUrl = value;
            }
        }
        
        /// <summary>
        /// This method is called to apply credentials to a service request before the request is made.
        /// </summary>
        /// <param name="request">The request.</param>
        internal override void PrepareWebRequest(IEwsHttpWebRequest request)
        {
            if ((this.EwsUrl == null) || (this.EwsUrl != request.RequestUri))
            {
                this.IsAuthenticated = false;
                this.MakeTokenRequestToWindowsLive(request.RequestUri);

                this.IsAuthenticated = true;
                this.EwsUrl = request.RequestUri;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="WindowsLiveCredentials"/> has been authenticated.
        /// </summary>
        /// <value><c>true</c> if authenticated; otherwise, <c>false</c>.</value>
        public bool IsAuthenticated
        {
            get { return this.isAuthenticated; }
            internal set { this.isAuthenticated = value; }
        }

        /// <summary>
        /// Function that sends the token request to Windows Live.
        /// </summary>
        /// <param name="uriForTokenEndpointReference">The Uri to use for the endpoint reference for our token</param>
        /// <returns>Response to token request.</returns>
        private HttpWebResponse EmitTokenRequest(Uri uriForTokenEndpointReference)
        {
            const string TokenRequest =
                "<?xml version='1.0' encoding='UTF-8'?>" +
                "<s:Envelope xmlns:s='http://www.w3.org/2003/05/soap-envelope' " +
                "            xmlns:wsse='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd' " +
                "            xmlns:saml='urn:oasis:names:tc:SAML:1.0:assertion' " +
                "            xmlns:wsp='http://schemas.xmlsoap.org/ws/2004/09/policy' " +
                "            xmlns:wsu='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd' " +
                "            xmlns:wsa='http://www.w3.org/2005/08/addressing' " +
                "            xmlns:wssc='http://schemas.xmlsoap.org/ws/2005/02/sc' " +
                "            xmlns:wst='http://schemas.xmlsoap.org/ws/2005/02/trust' " +
                "            xmlns:ps='http://schemas.microsoft.com/Passport/SoapServices/PPCRL'>" +
                "  <s:Header>" +
                "    <wsa:Action s:mustUnderstand='1'>http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action>" +
                "    <wsa:To s:mustUnderstand='1'>{0}</wsa:To>" +
                "    <ps:AuthInfo Id='PPAuthInfo'>" +
                "      <ps:HostingApp>{{63f179af-8bcd-49a0-a3e5-1154c02df090}}</ps:HostingApp>" + //// NOTE: I generated a new GUID for the EWS API
                "      <ps:BinaryVersion>5</ps:BinaryVersion>" +
                "      <ps:UIVersion>1</ps:UIVersion>" +
                "      <ps:Cookies></ps:Cookies>" +
                "      <ps:RequestParams>AQAAAAIAAABsYwQAAAAxMDMz</ps:RequestParams>" +
                "    </ps:AuthInfo>" +
                "    <wsse:Security>" +
                "      <wsse:UsernameToken wsu:Id='user'>" +
                "        <wsse:Username>{1}</wsse:Username>" +
                "        <wsse:Password>{2}</wsse:Password>" +
                "      </wsse:UsernameToken>" +
                "      <wsu:Timestamp Id='Timestamp'>" +
                "        <wsu:Created>{3}</wsu:Created>" +
                "        <wsu:Expires>{4}</wsu:Expires>" +
                "      </wsu:Timestamp>" +
                "    </wsse:Security>" +
                "  </s:Header>" +
                "  <s:Body>" +
                "    <ps:RequestMultipleSecurityTokens Id='RSTS'>" +
                "      <wst:RequestSecurityToken Id='RST0'>" +
                "        <wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType>" +
                "        <wsp:AppliesTo>" +
                "          <wsa:EndpointReference>" +
                "            <wsa:Address>http://Passport.NET/tb</wsa:Address>" +
                "          </wsa:EndpointReference>" +
                "        </wsp:AppliesTo>" +
                "      </wst:RequestSecurityToken>" +
                "      <wst:RequestSecurityToken Id='RST1'>" +
                "        <wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType>" +
                "        <wsp:AppliesTo>" +
                "          <wsa:EndpointReference>" +
                "            <wsa:Address>{5}</wsa:Address>" +
                "          </wsa:EndpointReference>" +
                "        </wsp:AppliesTo>" +
                "        <wsp:PolicyReference URI='LBI_FED_SSL'></wsp:PolicyReference>" +
                "      </wst:RequestSecurityToken>" +
                "    </ps:RequestMultipleSecurityTokens>" +
                "  </s:Body>" +
                "</s:Envelope>";
            
            // Create a security timestamp valid for 5 minutes to send with the request.
            DateTime now = DateTime.UtcNow;
            SecurityTimestamp securityTimestamp = new SecurityTimestamp(now, now.AddMinutes(5), "Timestamp");

            // Format the request string to send to the server, filling in all the bits.
            string requestToSend = String.Format(
                TokenRequest,
                this.windowsLiveUrl,
                this.windowsLiveId,
                this.password,
                securityTimestamp.GetCreationTimeChars(),
                securityTimestamp.GetExpiryTimeChars(),
                uriForTokenEndpointReference.ToString());

            // Create and send the request.
            HttpWebRequest webRequest = (HttpWebRequest) HttpWebRequest.Create(this.windowsLiveUrl);
            
            webRequest.Method = "POST";
            webRequest.ContentType = "text/xml; charset=utf-8";
            byte[] requestBytes = Encoding.UTF8.GetBytes(requestToSend);
            webRequest.ContentLength = requestBytes.Length;

            // NOTE: We're not tracing the request to Windows Live here because it has the user name and
            // password in it.
            using (Stream requestStream = webRequest.GetRequestStream())
            {
                requestStream.Write(requestBytes, 0, requestBytes.Length);
            }
            
            return (HttpWebResponse)webRequest.GetResponse();
        }

        /// <summary>
        /// Traces the response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="memoryStream">The response content in a MemoryStream.</param>
        private void TraceResponse(HttpWebResponse response, MemoryStream memoryStream)
        {
            EwsUtilities.Assert(
                memoryStream != null,
                "WindowsLiveCredentials.TraceResponse",
                "memoryStream cannot be null");
            
            if (!this.TraceEnabled)
            {
                return;
            }
            
            if (!string.IsNullOrEmpty(response.ContentType) && 
                (response.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) ||
                 response.ContentType.StartsWith("application/soap", StringComparison.OrdinalIgnoreCase)))
            {
                this.traceListener.Trace(
                    "WindowsLiveResponse",
                    EwsUtilities.FormatLogMessageWithXmlContent("WindowsLiveResponse", memoryStream));
            }
            else
            {
                this.traceListener.Trace(
                    "WindowsLiveResponse",
                    "Non-textual response");
            }
        }

        private void TraceWebException(WebException e)
        {
            // If there wasn't a response, there's nothing to trace.
            if (e.Response == null)
            {
                if (this.TraceEnabled)
                {
                    string logMessage = string.Format(
                        "Exception Received when sending Windows Live token request: {0}",
                        e);
                    this.traceListener.Trace("WindowsLiveResponse", logMessage);
                }
                return;
            }

            // If tracing is enabled, we read the entire response into a MemoryStream so that we
            // can pass it along to the ITraceListener. Then we parse the response from the 
            // MemoryStream.
            if (this.TraceEnabled)
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Stream responseStream = e.Response.GetResponseStream())
                    {
                        // Copy response to in-memory stream and reset position to start.
                        EwsUtilities.CopyStream(responseStream, memoryStream);
                        memoryStream.Position = 0;
                    }

                    this.TraceResponse((HttpWebResponse) e.Response, memoryStream);
                }
            }
        }
        
        /// <summary>
        /// Makes a request to Windows Live to get a token.
        /// </summary>
        /// <param name="uriForTokenEndpointReference">URL where token is to be used</param>
        private void MakeTokenRequestToWindowsLive(Uri uriForTokenEndpointReference)
        {
            // Post the request to Windows Live and load the response into an EwsXmlReader for
            // processing.
            HttpWebResponse response;

            try
            {
                response = this.EmitTokenRequest(uriForTokenEndpointReference);
            }
            catch (WebException e)
            {
                if (e.Status == WebExceptionStatus.ProtocolError && e.Response != null)
                {
                    this.TraceWebException(e);
                }
                else
                {
                    if (this.TraceEnabled)
                    {
                        string traceString = string.Format(
                            "Error occurred sending request - status was {0}, exception {1}",
                            e.Status,
                            e);
                        this.traceListener.Trace(
                            "WindowsLiveCredentials",
                            traceString);
                    }
                }
                
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }

            try
            {
                this.ProcessTokenResponse(response);
            }
            catch (WebException e)
            {
                if (this.TraceEnabled)
                {
                    string traceString = string.Format(
                        "Error occurred sending request - status was {0}, exception {1}",
                        e.Status,
                        e);
                    this.traceListener.Trace(
                        "WindowsLiveCredentials",
                        traceString);
                }
                
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
            }
        }

        /// <summary>
        /// Function that parses the SOAP headers from the response to the RST to Windows Live.
        /// </summary>
        /// <param name="rstResponse">The Windows Live response, positioned at the beginning of the SOAP headers.</param>
        private void ReadWindowsLiveRSTResponseHeaders(EwsXmlReader rstResponse)
        {
            // Read the beginning of the SOAP header, then go looking for the Passport SOAP fault section...
            rstResponse.ReadStartElement(
                WindowsLiveSoapNamespacePrefix,
                XmlElementNames.SOAPHeaderElementName);

            // Attempt to read to the psf:pp element - if at the end of the ReadToDescendant call we're at the
            // end element for the SOAP headers, we didn't find it.
            rstResponse.ReadToDescendant(XmlNamespace.PassportSoapFault, PpElementName);
            if (rstResponse.IsEndElement(WindowsLiveSoapNamespacePrefix, XmlElementNames.SOAPHeaderElementName))
            {
                // We didn't find the psf:pp element - without that, we don't know what happened -
                // something went wrong.  Trace and throw.
                if (this.TraceEnabled)
                {
                    this.traceListener.Trace(
                        "WindowsLiveResponse",
                        "Could not find Passport SOAP fault information in Windows Live response");
                }
                
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, PpElementName));
            }
            
            // Now that we've found the psf:pp element, look for the 'reqstatus' element under it.  If after
            // the ReadToDescendant call we're at the end element for the psf:pp element, we didn't find it.
            rstResponse.ReadToDescendant(XmlNamespace.PassportSoapFault, ReqstatusElementName);
            if (rstResponse.IsEndElement(XmlNamespace.PassportSoapFault, PpElementName))
            {
                // We didn't find the "reqstatus" element - without that, we don't know what happened -
                // something went wrong.  Trace and throw.
                if (this.TraceEnabled)
                {
                    this.traceListener.Trace(
                        "WindowsLiveResponse",
                        "Could not find reqstatus element in Passport SOAP fault information in Windows Live response");
                }
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ReqstatusElementName));
            }
            
            // Now that we've found the reqstatus element, get its value.
            string reqstatus = rstResponse.ReadElementValue();

            // Read to body tag in both success and failure cases, 
            // since we need to trace the fault response in failure cases
            while (!rstResponse.IsEndElement(WindowsLiveSoapNamespacePrefix, XmlElementNames.SOAPHeaderElementName))
            {
                rstResponse.Read();
            }
            
            if (!string.Equals(reqstatus, SuccessfulReqstatus))
            {
                // Our request status was non-zero - something went wrong.  Trace and throw.
                if (this.TraceEnabled)
                {
                    string logMessage = string.Format(
                        "Received status {0} from Windows Live instead of {1}.",
                        reqstatus,
                        SuccessfulReqstatus);
                    this.traceListener.Trace("WindowsLiveResponse", logMessage);

                    rstResponse.ReadStartElement(
                        WindowsLiveSoapNamespacePrefix,
                        XmlElementNames.SOAPBodyElementName);

                    // Trace Fault Information
                    this.traceListener.Trace(
                           "WindowsLiveResponse",
                           string.Format(
                           "Windows Live reported Fault : {0}",
                           rstResponse.ReadInnerXml()));
                }               
                
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ReqstatusElementName + ": " + reqstatus));
            }
        }

        /// <summary>
        /// Function that parses the RSTR from Windows Live and pulls out all the important pieces
        /// of data from it.
        /// </summary>
        /// <param name="rstResponse">The RSTR, positioned at the beginning of the SOAP body.</param>
        private void ParseWindowsLiveRSTResponseBody(EwsXmlReader rstResponse)
        {
            // Read the WS-Trust RequestSecurityTokenResponseCollection node.
            rstResponse.ReadStartElement(
                XmlNamespace.WSTrustFebruary2005,
                RequestSecurityTokenResponseCollectionElementName);

            // Skip the first token - our interest is in the second token (the service token).
            rstResponse.SkipElement(
                XmlNamespace.WSTrustFebruary2005,
                RequestSecurityTokenResponseElementName);

            // Now process the second token.
            rstResponse.ReadStartElement(
                XmlNamespace.WSTrustFebruary2005,
                RequestSecurityTokenResponseElementName);

            while (!rstResponse.IsEndElement(
                       XmlNamespace.WSTrustFebruary2005,
                       RequestSecurityTokenResponseElementName))
            {
                // Watch for the EncryptedData element - when we find it, parse out the appropriate bits of data.
                //
                // Also watch for the "pp" element in the Passport SOAP fault namespace, which indicates that
                // something went wrong with the token request.  If we find it, trace and throw accordingly.
                if (rstResponse.IsStartElement() &&
                    (rstResponse.LocalName == EncryptedDataElementName) &&
                    (rstResponse.NamespaceUri == XmlEncNamespace))
                {
                    this.SecurityToken = rstResponse.ReadOuterXml();
                }
                else if (rstResponse.IsStartElement(XmlNamespace.PassportSoapFault, PpElementName))
                {
                    if (this.TraceEnabled)
                    {
                        string logMessage = string.Format(
                            "Windows Live reported an error retrieving the token - {0}",
                            rstResponse.ReadOuterXml());
                        this.traceListener.Trace("WindowsLiveResponse", logMessage);
                    }
                    throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, EncryptedDataElementName));
                }

                // Move to the next bit of data...
                rstResponse.Read();
            }

            // If we didn't find the token, throw.
            if (this.SecurityToken == null)
            {
                if (this.TraceEnabled)
                {
                    string logMessage = string.Format(
                        "Did not find all required parts of the Windows Live response - " +
                        "Security Token - {0}",
                        (this.SecurityToken == null) ? "NOT FOUND" : "found");
                    this.traceListener.Trace("WindowsLiveResponse", logMessage);
                }
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, "No security token found."));
            }
            
            // Read past the RequestSecurityTokenResponseCollection end element.
            rstResponse.Read();
        }
        
        /// <summary>
        /// Grabs the issued token information out of a response from Windows Live.
        /// </summary>
        /// <param name="response">The token response</param>
        private void ProcessTokenResponse(HttpWebResponse response)
        {
            // NOTE: We're not tracing responses here because they contain the actual token information
            // from Windows Live.    
            using (Stream responseStream = response.GetResponseStream())
            {
                // Always start fresh (nulls in all the data we're going to fill in).
                this.SecurityToken = null;

                EwsXmlReader rstResponse = new EwsXmlReader(responseStream);

                rstResponse.Read(XmlNodeType.XmlDeclaration);
                rstResponse.ReadStartElement(
                    WindowsLiveSoapNamespacePrefix,
                    XmlElementNames.SOAPEnvelopeElementName);

                // Process the SOAP headers from the response.
                this.ReadWindowsLiveRSTResponseHeaders(rstResponse);

                rstResponse.ReadStartElement(
                    WindowsLiveSoapNamespacePrefix,
                    XmlElementNames.SOAPBodyElementName);

                // Process the SOAP body from the response.
                this.ParseWindowsLiveRSTResponseBody(rstResponse);
            }
        }
    }
}