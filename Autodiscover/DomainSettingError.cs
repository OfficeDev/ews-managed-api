// ---------------------------------------------------------------------------
// <copyright file="DomainSettingError.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the DomainSettingError class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an error from a GetDomainSettings request.
    /// </summary>
    public sealed class DomainSettingError
    {
        private AutodiscoverErrorCode errorCode;
        private string errorMessage;
        private string settingName;

        /// <summary>
        /// Initializes a new instance of the <see cref="DomainSettingError"/> class.
        /// </summary>
        internal DomainSettingError()
        {
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsXmlReader reader)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.ErrorCode:
                            this.errorCode = reader.ReadElementValue<AutodiscoverErrorCode>();
                            break;
                        case XmlElementNames.ErrorMessage:
                            this.errorMessage = reader.ReadElementValue();
                            break;
                        case XmlElementNames.SettingName:
                            this.settingName = reader.ReadElementValue();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSettingError));
        }

        /// <summary>
        /// Gets the error code.
        /// </summary>
        /// <value>The error code.</value>
        public AutodiscoverErrorCode ErrorCode
        {
            get { return this.errorCode; }
        }

        /// <summary>
        /// Gets the error message.
        /// </summary>
        /// <value>The error message.</value>
        public string ErrorMessage
        {
            get { return this.errorMessage; }
        }

        /// <summary>
        /// Gets the name of the setting.
        /// </summary>
        /// <value>The name of the setting.</value>
        public string SettingName
        {
            get { return this.settingName; }
        }
    }
}
