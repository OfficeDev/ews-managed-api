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
// <summary>Defines the SetEncryptionConfigurationRequest class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents a SetEncryptionConfiguration request.
    /// </summary>
    internal sealed class SetEncryptionConfigurationRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// The base64 encoding of the image
        /// </summary>
        private readonly string imageBase64;

        /// <summary>
        /// The email text
        /// </summary>
        private readonly string emailText;

        /// <summary>
        /// The portal text
        /// </summary>
        private readonly string portalText;

        /// <summary>
        /// The disclaimer text
        /// </summary>
        private readonly string disclaimerText;

        /// <summary>
        /// If OTP is enabled
        /// </summary>
        private readonly bool otpEnabled;

        /// <summary>
        /// The base64 encoding of the image
        /// </summary>
        public string ImageBase64
        {
            get { return this.imageBase64; }
        }

        /// <summary>
        /// The EmailText
        /// </summary>
        public string EmailText
        {
            get { return this.emailText; }
        }

        /// <summary>
        /// The PortalText
        /// </summary>
        public string PortalText
        {
            get { return this.portalText; }
        }

        /// <summary>
        /// The DisclaimerText
        /// </summary>
        public string DisclaimerText
        {
            get { return this.disclaimerText; }
        }

        /// <summary>
        /// If OTP is enabled
        /// </summary>
        public bool OTPEnabled
        {
            get { return this.otpEnabled; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SetEncryptionConfigurationRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <param name="imageBase64">The base64 encoding of the image</param>
        /// <param name="emailText">The email text</param>
        /// <param name="portalText">The portal text</param>
        /// <param name="disclaimerText">The disclaimer text</param>
        /// <param name="otpEnabled">If OTP is enabled</param>
        internal SetEncryptionConfigurationRequest(
            ExchangeService service,
            string imageBase64,
            string emailText,
            string portalText,
            string disclaimerText,
            bool otpEnabled)
                : base(service)
        {
            this.emailText = emailText;
            this.portalText = portalText;
            this.imageBase64 = imageBase64;
            this.disclaimerText = disclaimerText;
            this.otpEnabled = otpEnabled;
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.SetEncryptionConfigurationRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationImageBase64, this.ImageBase64);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationEmailText, this.EmailText);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationPortalText, this.PortalText);
            writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationDisclaimerText, this.disclaimerText);

            // TODO: Uncomment after server schema change have been deployed to all machines
            // writer.WriteElementValue(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationOTPEnabled, this.otpEnabled);
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.SetEncryptionConfigurationResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            SetEncryptionConfigurationResponse response = new SetEncryptionConfigurationResponse();
            response.LoadFromXml(reader, GetResponseXmlElementName());
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal ServiceResponse Execute()
        {
            SetEncryptionConfigurationResponse serviceResponse = (SetEncryptionConfigurationResponse)this.InternalExecute();
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}