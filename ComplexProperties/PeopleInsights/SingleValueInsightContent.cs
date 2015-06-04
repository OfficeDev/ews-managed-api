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
    /// Represents the SingleValueInsightContent.
    /// </summary>
    public sealed class SingleValueInsightContent : ComplexProperty
    {
        /// <summary>
        /// Gets the Item
        /// </summary>
        public InsightValue Item { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.ReadAttributeValue("xsi:type"))
            {
                case XmlElementNames.StringInsightValue:
                    this.Item = new StringInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.ProfileInsightValue:
                    this.Item = new ProfileInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.JobInsightValue:
                    this.Item = new JobInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.UserProfilePicture:
                    this.Item = new UserProfilePicture();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.EducationInsightValue:
                    this.Item = new EducationInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.SkillInsightValue:
                    this.Item = new SkillInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.DelveDocument:
                    this.Item = new DelveDocument();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.CompanyInsightValue:
                    this.Item = new CompanyInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.ComputedInsightValue:
                    this.Item = new ComputedInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.OutOfOfficeInsightValue:
                    this.Item = new OutOfOfficeInsightValue();
                    this.Item.LoadFromXml(reader, reader.LocalName);
                    break;
                default:
                    return false;
            }

            return true;
        }
    }   
}