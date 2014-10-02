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
// <summary>Defines the GetEncryptionConfigurationResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a GetEncryptionConfiguration operation.
    /// </summary>
    public sealed class GetEncryptionConfigurationResponse : ServiceResponse
    {
        /// <summary>
        /// The base64 encoding of the image
        /// </summary>
        private string imageBase64;

        /// <summary>
        /// The email text
        /// </summary>
        private string emailText;

        /// <summary>
        /// The portal text
        /// </summary>
        private string portalText;

        /// <summary>
        /// The disclaimer text
        /// </summary>
        private string disclaimerText;

        /// <summary>
        /// If OTP is enabled
        /// </summary>
        private bool otpEnabled;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetEncryptionConfigurationResponse"/> class.
        /// </summary>
        internal GetEncryptionConfigurationResponse()
            : base()
        {
        }

        /// <summary>
        /// The base64 encoding of the Image
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
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.imageBase64 = reader.ReadElementValue<string>(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationImageBase64);
            this.emailText = reader.ReadElementValue<string>(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationEmailText);
            this.portalText = reader.ReadElementValue<string>(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationPortalText);
            this.disclaimerText = reader.ReadElementValue<string>(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationDisclaimerText);

            // TODO: Remove the try/catch after both client & server have been deployed to all machines
            try
            {
                this.otpEnabled = reader.ReadElementValue<bool>(XmlNamespace.Messages, XmlElementNames.EncryptionConfigurationOTPEnabled);
            }
            catch (ServiceXmlDeserializationException)
            {
                this.otpEnabled = true;
            }
        }
    }
}
