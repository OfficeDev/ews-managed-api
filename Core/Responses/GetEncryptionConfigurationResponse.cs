// ---------------------------------------------------------------------------
// <copyright file="GetEncryptionConfigurationResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
