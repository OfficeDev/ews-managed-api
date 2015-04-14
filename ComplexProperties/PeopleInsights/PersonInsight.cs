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
        }

        /// <summary>
        /// Gets the InsightGroupType
        /// </summary>
        public InsightGroupType InsightGroupType { get; internal set; }

        /// <summary>
        /// Gets the InsightType
        /// </summary>
        public InsightType InsightType { get; internal set; }

        /// <summary>
        /// Gets the Rank
        /// </summary>
        public double Rank { get; internal set; }

        /// <summary>
        /// Gets the Content
        /// </summary>
        public ComplexProperty Content { get; internal set; }

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
                    case XmlElementNames.InsightGroupType:
                        this.InsightGroupType = reader.ReadElementValue<InsightGroupType>();
                        break;
                    case XmlElementNames.InsightType:
                        this.InsightType = reader.ReadElementValue<InsightType>();
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
                    default:
                        return false;
                }

                return true;
            }
        }
    }
}