// ---------------------------------------------------------------------------
// <copyright file="UserSettingError.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UserSettingError class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents an error from a GetUserSettings request.
    /// </summary>
    public sealed class UserSettingError
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserSettingError"/> class.
        /// </summary>
        internal UserSettingError()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserSettingError"/> class.
        /// </summary>
        /// <param name="errorCode">The error code.</param>
        /// <param name="errorMessage">The error message.</param>
        /// <param name="settingName">Name of the setting.</param>
        internal UserSettingError(
            AutodiscoverErrorCode errorCode,
            string errorMessage,
            string settingName)
        {
            this.ErrorCode = errorCode;
            this.ErrorMessage = errorMessage;
            this.SettingName = settingName;
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
                            this.ErrorCode = reader.ReadElementValue<AutodiscoverErrorCode>();
                            break;
                        case XmlElementNames.ErrorMessage:
                            this.ErrorMessage = reader.ReadElementValue();
                            break;
                        case XmlElementNames.SettingName:
                            this.SettingName = reader.ReadElementValue();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.UserSettingError));
        }

        /// <summary>
        /// Gets the error code.
        /// </summary>
        /// <value>The error code.</value>
        public AutodiscoverErrorCode ErrorCode
        {
            get; internal set;
        }

        /// <summary>
        /// Gets the error message.
        /// </summary>
        /// <value>The error message.</value>
        public string ErrorMessage
        {
            get; internal set;
        }

        /// <summary>
        /// Gets the name of the setting.
        /// </summary>
        /// <value>The name of the setting.</value>
        public string SettingName
        {
            get; internal set;
        }
    }
}
