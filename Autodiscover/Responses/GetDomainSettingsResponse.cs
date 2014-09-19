// ---------------------------------------------------------------------------
// <copyright file="GetDomainSettingsResponse.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the GetDomainSettingsResponse class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Text;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents the response to a GetDomainSettings call for an individual domain.
    /// </summary>
    public sealed class GetDomainSettingsResponse : AutodiscoverResponse
    {
        private string domain;
        private string redirectTarget;
        private Dictionary<DomainSettingName, object> settings;
        private Collection<DomainSettingError> domainSettingErrors;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetDomainSettingsResponse"/> class.
        /// </summary>
        public GetDomainSettingsResponse()
            : base()
        {
            this.domain = string.Empty;
            this.settings = new Dictionary<DomainSettingName, object>();
            this.domainSettingErrors = new Collection<DomainSettingError>();
        }

        /// <summary>
        /// Gets the domain this response applies to.
        /// </summary>
        public string Domain
        {
            get { return this.domain; }
            internal set { this.domain = value; }
        }

        /// <summary>
        /// Gets the redirectionTarget (URL or email address)
        /// </summary>
        public string RedirectTarget
        {
            get { return this.redirectTarget; }
        }

        /// <summary>
        /// Gets the requested settings for the domain.
        /// </summary>
        public IDictionary<DomainSettingName, object> Settings
        {
            get { return this.settings; }
        }

        /// <summary>
        /// Gets error information for settings that could not be returned.
        /// </summary>
        public Collection<DomainSettingError> DomainSettingErrors
        {
            get { return this.domainSettingErrors; }
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
                            this.redirectTarget = reader.ReadElementValue();
                            break;
                        case XmlElementNames.DomainSettingErrors:
                            this.LoadDomainSettingErrorsFromXml(reader);
                            break;
                        case XmlElementNames.DomainSettings:
                            this.LoadDomainSettingsFromXml(reader);
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
        internal void LoadDomainSettingsFromXml(EwsXmlReader reader)
        {
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.DomainSetting))
                    {
                        string settingClass = reader.ReadAttributeValue(XmlNamespace.XmlSchemaInstance, XmlAttributeNames.Type);

                        switch (settingClass)
                        {
                            case XmlElementNames.DomainStringSetting:
                                this.ReadSettingFromXml(reader);
                                break;

                            default:
                                EwsUtilities.Assert(
                                    false,
                                    "GetDomainSettingsResponse.LoadDomainSettingsFromXml",
                                    string.Format("Invalid setting class '{0}' returned", settingClass));
                                break;
                        }
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSettings));
            }
        }

        /// <summary>
        /// Reads domain setting from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadSettingFromXml(EwsXmlReader reader)
        {
            DomainSettingName? name = null;
            object value = null;

            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case XmlElementNames.Name:
                            name = reader.ReadElementValue<DomainSettingName>();
                            break;
                        case XmlElementNames.Value:
                            value = reader.ReadElementValue();
                            break;
                    }
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSetting));

            EwsUtilities.Assert(
                name.HasValue,
                "GetDomainSettingsResponse.ReadSettingFromXml",
                "Missing name element in domain setting");

            this.settings.Add(name.Value, value);
        }

        /// <summary>
        /// Loads the domain setting errors.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void LoadDomainSettingErrorsFromXml(EwsXmlReader reader)
        {
            if (!reader.IsEmptyElement)
            {
                do
                {
                    reader.Read();

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.DomainSettingError))
                    {
                        DomainSettingError error = new DomainSettingError();
                        error.LoadFromXml(reader);
                        domainSettingErrors.Add(error);
                    }
                }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSettingErrors));
            }
        }
    }
}
