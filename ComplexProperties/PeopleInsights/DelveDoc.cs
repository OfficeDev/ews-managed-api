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
    /// Represents the DelveDoc.
    /// </summary>
    public sealed class DelveDoc : InsightValue
    {
        private double rank;
        private string author;
        private string created;
        private string lastModifiedTime;
        private string defaultEncodingUrl;
        private string fileType;
        private string title;

        /// <summary>
        /// Gets the Rank
        /// </summary>
        public double Rank
        {
            get
            {
                return this.rank;
            }

            set
            {
                this.SetFieldValue<double>(ref this.rank, value);
            }
        }

        /// <summary>
        /// Gets the Author
        /// </summary>
        public string Author
        {
            get
            {
                return this.author;
            }

            set
            {
                this.SetFieldValue<string>(ref this.author, value);
            }
        }

        /// <summary>
        /// Gets the Created
        /// </summary>
        public string Created
        {
            get
            {
                return this.created;
            }

            set
            {
                this.SetFieldValue<string>(ref this.created, value);
            }
        }

        /// <summary>
        /// Gets the LastModifiedTime
        /// </summary>
        public string LastModifiedTime
        {
            get
            {
                return this.lastModifiedTime;
            }

            set
            {
                this.SetFieldValue<string>(ref this.lastModifiedTime, value);
            }
        }

        /// <summary>
        /// Gets the DefaultEncodingURL
        /// </summary>
        public string DefaultEncodingURL
        {
            get
            {
                return this.defaultEncodingUrl;
            }

            set
            {
                this.SetFieldValue<string>(ref this.defaultEncodingUrl, value);
            }
        }

        /// <summary>
        /// Gets the FileType
        /// </summary>
        public string FileType
        {
            get
            {
                return this.fileType;
            }

            set
            {
                this.SetFieldValue<string>(ref this.fileType, value);
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
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.InsightSource:
                    this.InsightSource = reader.ReadElementValue<InsightSourceType>();
                    break;
                case XmlElementNames.UpdatedUtcTicks:
                    this.UpdatedUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.Rank:
                    this.Rank = reader.ReadElementValue<double>();
                    break;
                case XmlElementNames.Author:
                    this.Author = reader.ReadElementValue();
                    break;
                case XmlElementNames.Created:
                    this.Created = reader.ReadElementValue();
                    break;
                case XmlElementNames.LastModifiedTime:
                    this.LastModifiedTime = reader.ReadElementValue();
                    break;
                case XmlElementNames.DefaultEncodingURL:
                    this.DefaultEncodingURL = reader.ReadElementValue();
                    break;
                case XmlElementNames.FileType:
                    this.FileType = reader.ReadElementValue();
                    break;
                case XmlElementNames.Title:
                    this.Title = reader.ReadElementValue();
                    break;
                default:
                    return false;
            }

            return true;
        }
    }
}