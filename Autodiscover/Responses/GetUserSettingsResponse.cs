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

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the response to a GetUsersSettings call for an individual user.
    /// </summary>
    public sealed class GetUserSettingsResponse : AutodiscoverResponse
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetUserSettingsResponse"/> class.
        /// </summary>
        public GetUserSettingsResponse()
            : base()
        {
            this.SmtpAddress = string.Empty;
            this.Settings = new Dictionary<UserSettingName, object>();
            this.UserSettingErrors = new Collection<UserSettingError>();
        }

        /// <summary>
        /// Tries the get the user setting value.
        /// </summary>
        /// <typeparam name="T">Type of user setting.</typeparam>
        /// <param name="setting">The setting.</param>
        /// <param name="value">The setting value.</param>
        /// <returns>True if setting was available.</returns>
        public bool TryGetSettingValue<T>(UserSettingName setting, out T value)
        {
            object objValue;
            if (this.Settings.TryGetValue(setting, out objValue))
            {
                value = (T)objValue;
                return true;
            }
            else
            {
                value = default(T);
                return false;
            }
        }

        /// <summary>
        /// Gets the SMTP address this response applies to.
        /// </summary>
        public string SmtpAddress
        {
            get; internal set;
        }

        /// <summary>
        /// Gets the redirectionTarget (URL or email address)
        /// </summary>
        public string RedirectTarget
        {
            get; internal set;
        }

        /// <summary>
        /// Gets the requested settings for the user.
        /// </summary>
        public IDictionary<UserSettingName, object> Settings
        {
            get; internal set;
        }

        /// <summary>
        /// Gets error information for settings that could not be returned.
        /// </summary>
        public Collection<UserSettingError> UserSettingErrors
        {
            get; internal set;
        }

        /// <summary>
        /// Loads response from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="endElementName">End element name.</param>
        internal override void LoadFromXml(EwsXmlReader reader, string endElementName)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.RedirectTarget:
                            this.RedirectTarget = reader.ReadElementValue();
                            break;
                        case XmlElementNames.UserSettingErrors:
                            this.LoadUserSettingErrorsFromXml(reader);
                            break;
                        case XmlElementNames.UserSettings:
                            this.LoadUserSettingsFromXml(reader);
                            break;
                        default:
                            base.LoadFromXml(reader, endElementName);
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, endElementName));
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadUserSettingsFromXml(EwsXmlReader reader)
        {
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.UserSetting))
                    {
                        string settingClass = reader.ReadAttributeValue(XmlNamespace.XmlSchemaInstance, XmlAttributeNames.Type);

                        switch (settingClass)
                        {
                            case XmlElementNames.StringSetting:
                            case XmlElementNames.WebClientUrlCollectionSetting:
                            case XmlElementNames.AlternateMailboxCollectionSetting:
                            case XmlElementNames.ProtocolConnectionCollectionSetting:
                            case XmlElementNames.DocumentSharingLocationCollectionSetting:
                                this.ReadSettingFromXml(reader);
                                break;

                            default:
                                EwsUtilities.Assert(
                                    false,
                                    "GetUserSettingsResponse.LoadUserSettingsFromXml",
                                    string.Format("Invalid setting class '{0}' returned", settingClass));
                                break;
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.UserSettings));
            }
        }

        /// <summary>
        /// Reads user setting from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadSettingFromXml(EwsXmlReader reader)
        {
            string name = null;
            object value = null;

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.Name:
                            name = reader.ReadElementValue<string>();
                            break;
                        case XmlElementNames.Value:
                            value = reader.ReadElementValue();
                            break;
                        case XmlElementNames.WebClientUrls:
                            value = WebClientUrlCollection.LoadFromXml(reader);
                            break;
                        case XmlElementNames.ProtocolConnections:
                            value = ProtocolConnectionCollection.LoadFromXml(reader);
                            break;
                        case XmlElementNames.AlternateMailboxes:
                            value = AlternateMailboxCollection.LoadFromXml(reader);
                            break;
                        case XmlElementNames.DocumentSharingLocations:
                            value = DocumentSharingLocationCollection.LoadFromXml(reader);
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.UserSetting));

            // EWS Managed API is broken with AutoDSvc endpoint in RedirectUrl scenario
            try
            {
                UserSettingName userSettingName = EwsUtilities.Parse<UserSettingName>(name);
                this.Settings.Add(userSettingName, value);
            }
            catch (ArgumentException)
            {
                // ignore unexpected UserSettingName in the response (due to the server-side bugs).
                // it'd be better if this is hooked into ITraceListener, but that is unavailable here.
                //
                // in case "name" is null, EwsUtilities.Parse throws ArgumentNullException 
                // (which derives from ArgumentException).
                //
                EwsUtilities.Assert(
                    false,
                    "GetUserSettingsResponse.ReadSettingFromXml",
                    "Unexpected or empty name element in user setting");
            }
        }

        /// <summary>
        /// Loads the user setting errors.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void LoadUserSettingErrorsFromXml(EwsXmlReader reader)
        {
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.UserSettingError))
                    {
                        UserSettingError error = new UserSettingError();
                        error.LoadFromXml(reader);
                        this.UserSettingErrors.Add(error);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.UserSettingErrors));
            }
        }
    }
}