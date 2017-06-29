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
    /// Represents the ProfileInsightValue.
    /// </summary>
    public sealed class ProfileInsightValue : InsightValue
    {
        private string fullName;
        private string firstName;
        private string lastName;
        private string emailAddress;
        private string avatar;
        private long joinedUtcTicks;
        private UserProfilePicture profilePicture;
        private string title;

        /// <summary>
        /// Gets the FullName
        /// </summary>
        public string FullName
        {
            get
            {
                return this.fullName;
            }
        }

        /// <summary>
        /// Gets the FirstName
        /// </summary>
        public string FirstName
        {
            get
            {
                return this.firstName;
            }
        }

        /// <summary>
        /// Gets the LastName
        /// </summary>
        public string LastName
        {
            get
            {
                return this.lastName;
            }
        }

        /// <summary>
        /// Gets the EmailAddress
        /// </summary>
        public string EmailAddress
        {
            get
            {
                return this.emailAddress;
            }
        }

        /// <summary>
        /// Gets the Avatar
        /// </summary>
        public string Avatar
        {
            get
            {
                return this.avatar;
            }
        }

        /// <summary>
        /// Gets the JoinedUtcTicks
        /// </summary>
        public long JoinedUtcTicks
        {
            get
            {
                return this.joinedUtcTicks;
            }
        }

        /// <summary>
        /// Gets the ProfilePicture
        /// </summary>
        public UserProfilePicture ProfilePicture
        {
            get
            {
                return this.profilePicture;
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
                case XmlElementNames.FullName:
                    this.fullName = reader.ReadElementValue();
                    break;
                case XmlElementNames.FirstName:
                    this.firstName = reader.ReadElementValue();
                    break;
                case XmlElementNames.LastName:
                    this.lastName = reader.ReadElementValue();
                    break;
                case XmlElementNames.EmailAddress:
                    this.emailAddress = reader.ReadElementValue();
                    break;
                case XmlElementNames.Avatar:
                    this.avatar = reader.ReadElementValue();
                    break;
                case XmlElementNames.JoinedUtcTicks:
                    this.joinedUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.ProfilePicture:
                    var picture = new UserProfilePicture();
                    picture.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ProfilePicture);
                    this.profilePicture = picture;
                    break;
                case XmlElementNames.Title:
                    this.title = reader.ReadElementValue();
                    break;
                default:
                    return false;
            }

            return true;
        }
    }
}