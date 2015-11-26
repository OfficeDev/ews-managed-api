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
    using System.Collections.Generic;
    using System.Xml;

    /// <summary>
    /// Represents the JobInsightValue.
    /// </summary>
    public sealed class JobInsightValue : InsightValue
    {
        private string company;
        private string companyDescription;
        private string companyTicker;
        private string companyLogoUrl;
        private string companyWebsiteUrl;
        private string companyLinkedInUrl;
        private string title;
        private long startUtcTicks;
        private long endUtcTicks;

        /// <summary>
        /// Gets the Company
        /// </summary>
        public string Company
        {
            get
            {
                return this.company;
            }

            set
            {
                this.SetFieldValue<string>(ref this.company, value);
            }
        }

        /// <summary>
        /// Gets the CompanyDescription
        /// </summary>
        public string CompanyDescription
        {
            get
            {
                return this.companyDescription;
            }

            set
            {
                this.SetFieldValue<string>(ref this.companyDescription, value);
            }
        }

        /// <summary>
        /// Gets the CompanyTicker
        /// </summary>
        public string CompanyTicker
        {
            get
            {
                return this.companyTicker;
            }

            set
            {
                this.SetFieldValue<string>(ref this.companyTicker, value);
            }
        }

        /// <summary>
        /// Gets the CompanyLogoUrl
        /// </summary>
        public string CompanyLogoUrl
        {
            get
            {
                return this.companyLogoUrl;
            }

            set
            {
                this.SetFieldValue<string>(ref this.companyLogoUrl, value);
            }
        }

        /// <summary>
        /// Gets the CompanyWebsiteUrl
        /// </summary>
        public string CompanyWebsiteUrl
        {
            get
            {
                return this.companyWebsiteUrl;
            }

            set
            {
                this.SetFieldValue<string>(ref this.companyWebsiteUrl, value);
            }
        }

        /// <summary>
        /// Gets the CompanyLinkedInUrl
        /// </summary>
        public string CompanyLinkedInUrl
        {
            get
            {
                return this.companyLinkedInUrl;
            }

            set
            {
                this.SetFieldValue<string>(ref this.companyLinkedInUrl, value);
            }
        }

        /// <summary>
        /// Gets the Title
        /// </summary>
        public string Title
        {
            get
            {
                return this.title;
            }

            set
            {
                this.SetFieldValue<string>(ref this.title, value);
            }
        }

        /// <summary>
        /// Gets the StartUtcTicks
        /// </summary>
        public long StartUtcTicks
        {
            get
            {
                return this.startUtcTicks;
            }

            set
            {
                this.SetFieldValue<long>(ref this.startUtcTicks, value);
            }
        }

        /// <summary>
        /// Gets the EndUtcTicks
        /// </summary>
        public long EndUtcTicks
        {
            get
            {
                return this.endUtcTicks;
            }

            set
            {
                this.SetFieldValue<long>(ref this.endUtcTicks, value);
            }
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.InsightSource:
                    this.InsightSource = reader.ReadElementValue<string>();
                    break;
                case XmlElementNames.UpdatedUtcTicks:
                    this.UpdatedUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.Company:
                    this.Company = reader.ReadElementValue();
                    break;
                case XmlElementNames.Title:
                    this.Title = reader.ReadElementValue();
                    break;
                case XmlElementNames.StartUtcTicks:
                    this.StartUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.EndUtcTicks:
                    this.EndUtcTicks = reader.ReadElementValue<long>();
                    break;
                default:
                    return false;
            }

            return true;
        }
    }
}