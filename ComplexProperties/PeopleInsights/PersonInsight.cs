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
    /// Represents the PersonInsight.
    /// </summary>
    public sealed class PersonInsight : ComplexProperty
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PersonInsight"/> class.
        /// </summary>
        public PersonInsight() : base()
        {
            this.ItemList = new InsightValueCollection();
        }

        /// <summary>
        /// Gets the string
        /// </summary>
        public string InsightType { get; internal set; }

        /// <summary>
        /// Gets the Rank
        /// </summary>
        public double Rank { get; internal set; }

        /// <summary>
        /// Gets the Content
        /// </summary>
        public ComplexProperty Content { get; internal set; }

        /// <summary>
        /// Gets the ItemList
        /// </summary>
        public InsightValueCollection ItemList { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            while (true)
            {
                switch (reader.LocalName)
                {
                    case XmlElementNames.InsightType:
                        this.InsightType = reader.ReadElementValue<string>();
                        break;
                    case XmlElementNames.Rank:
                        this.Rank = reader.ReadElementValue<double>();
                        break;
                    case XmlElementNames.Content:
                        var type = reader.ReadAttributeValue("xsi:type");
                        switch (type)
                        { 
                            case XmlElementNames.SingleValueInsightContent:
                                this.Content = new SingleValueInsightContent();
                                ((SingleValueInsightContent)this.Content).LoadFromXml(reader, reader.LocalName);
                                break;
                            case XmlElementNames.MultiValueInsightContent:
                                this.Content = new MultiValueInsightContent();
                                ((MultiValueInsightContent)this.Content).LoadFromXml(reader, reader.LocalName);
                                break;
                            default:
                                return false;
                        }
                        break;
                    case XmlElementNames.ItemList:
                        this.ReadItemList(reader);
                        break;
                    default:
                        return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Reads ItemList from XML
        /// </summary>
        /// <param name="reader">The reader.</param>        
        private void ReadItemList(EwsServiceXmlReader reader)
        {
            do
            {
                reader.Read();
                InsightValue item = null;

                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == XmlElementNames.Item)
                {
                    switch (reader.ReadAttributeValue("xsi:type"))
                    {
                        case XmlElementNames.StringInsightValue:
                            item = new StringInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.ProfileInsightValue:
                            item = new ProfileInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.JobInsightValue:
                            item = new JobInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.UserProfilePicture:
                            item = new UserProfilePicture();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.EducationInsightValue:
                            item = new EducationInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.SkillInsightValue:
                            item = new SkillInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.ComputedInsightValue:
                            item = new ComputedInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.MeetingInsightValue:
                            item = new MeetingInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.EmailInsightValue:
                            item = new EmailInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.DelveDocument:
                            item = new DelveDocument();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.CompanyInsightValue:
                            item = new CompanyInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                        case XmlElementNames.OutOfOfficeInsightValue:
                            item = new OutOfOfficeInsightValue();
                            item.LoadFromXml(reader, reader.LocalName);
                            this.ItemList.InternalAdd(item);
                            break;
                    }
                }
            } 
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.ItemList));
        }
    }
}