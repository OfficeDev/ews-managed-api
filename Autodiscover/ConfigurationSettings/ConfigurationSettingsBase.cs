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
    using System.ComponentModel;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the base class for configuration settings.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal abstract class ConfigurationSettingsBase
    {
        private AutodiscoverError error;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationSettingsBase"/> class.
        /// </summary>
        internal ConfigurationSettingsBase()
        {
        }

        /// <summary>
        /// Tries to read the current XML element.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True is the current element was read, false otherwise.</returns>
        internal virtual bool TryReadCurrentXmlElement(EwsXmlReader reader)
        {
            if (reader.LocalName == XmlElementNames.Error)
            {
                this.error = AutodiscoverError.Parse(reader);

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Loads the settings from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsXmlReader reader)
        {
            reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.Autodiscover);
            reader.ReadStartElement(XmlNamespace.NotSpecified, XmlElementNames.Response);

            do
            {
                reader.Read();

                if (reader.IsStartElement())
                {
                    if (!this.TryReadCurrentXmlElement(reader))
                    {
                        reader.SkipCurrentElement();
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.Response));

            reader.ReadEndElement(XmlNamespace.NotSpecified, XmlElementNames.Autodiscover);
        }

        /// <summary>
        /// Gets the namespace that defines the settings.
        /// </summary>
        /// <returns>The namespace that defines the settings</returns>
        internal abstract string GetNamespace();

        /// <summary>
        /// Makes this instance a redirection response.
        /// </summary>
        /// <param name="redirectUrl">The redirect URL.</param>
        internal abstract void MakeRedirectionResponse(Uri redirectUrl);

        /// <summary>
        /// Gets the type of the response.
        /// </summary>
        /// <value>The type of the response.</value>
        internal abstract AutodiscoverResponseType ResponseType { get; }

        /// <summary>
        /// Gets the redirect target.
        /// </summary>
        /// <value>The redirect target.</value>
        internal abstract string RedirectTarget { get; }

        /// <summary>
        /// Convert ConfigurationSettings to GetUserSettings response.
        /// </summary>
        /// <param name="smtpAddress">SMTP address.</param>
        /// <param name="requestedSettings">The requested settings.</param>
        /// <returns>GetUserSettingsResponse</returns>
        internal abstract GetUserSettingsResponse ConvertSettings(string smtpAddress, List<UserSettingName> requestedSettings);

        /// <summary>
        /// Gets the error.
        /// </summary>
        /// <value>The error.</value>
        internal AutodiscoverError Error
        {
            get { return this.error; }
        }
    }
}