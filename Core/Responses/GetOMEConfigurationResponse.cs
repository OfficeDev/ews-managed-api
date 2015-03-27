// ---------------------------------------------------------------------------
// <copyright file="GetOMEConfigurationResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetOMEConfigurationResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents the response to a GetOMEConfiguration operation.
    /// </summary>
    public sealed class GetOMEConfigurationResponse : ServiceResponse
    {
        /// <summary>
        /// The XML representation of EncryptionConfigurationData
        /// </summary>
        private string xml;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetOMEConfigurationResponse"/> class.
        /// </summary>
        internal GetOMEConfigurationResponse()
            : base()
        {
        }

        /// <summary>
        /// The XML representation of EncryptionConfigurationData
        /// </summary>
        public string Xml
        {
            get { return this.xml; }
        }

        /// <summary>
        /// Reads response elements from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadElementsFromXml(EwsServiceXmlReader reader)
        {
            base.ReadElementsFromXml(reader);

            this.xml = reader.ReadElementValue<string>(XmlNamespace.Messages, XmlElementNames.OMEConfigurationXml);
        }
    }
}
