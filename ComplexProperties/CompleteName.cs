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
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the complete name of a contact.
    /// </summary>
    public sealed class CompleteName : ComplexProperty
    {
        private string title;
        private string givenName;
        private string middleName;
        private string surname;
        private string suffix;
        private string initials;
        private string fullName;
        private string nickname;
        private string yomiGivenName;
        private string yomiSurname;

        #region Properties

        /// <summary>
        /// Gets the contact's title.
        /// </summary>
        public string Title
        {
            get { return this.title; }
        }

        /// <summary>
        /// Gets the given name (first name) of the contact.
        /// </summary>
        public string GivenName
        {
            get { return this.givenName; }
        }

        /// <summary>
        /// Gets the middle name of the contact.
        /// </summary>
        public string MiddleName
        {
            get { return this.middleName; }
        }

        /// <summary>
        /// Gets the surname (last name) of the contact.
        /// </summary>
        public string Surname
        {
            get { return this.surname; }
        }

        /// <summary>
        /// Gets the suffix of the contact.
        /// </summary>
        public string Suffix
        {
            get { return this.suffix; }
        }

        /// <summary>
        /// Gets the initials of the contact.
        /// </summary>
        public string Initials
        {
            get { return this.initials; }
        }

        /// <summary>
        /// Gets the full name of the contact.
        /// </summary>
        public string FullName
        {
            get { return this.fullName; }
        }

        /// <summary>
        /// Gets the nickname of the contact.
        /// </summary>
        public string NickName
        {
            get { return this.nickname; }
        }

        /// <summary>
        /// Gets the Yomi given name (first name) of the contact.
        /// </summary>
        public string YomiGivenName
        {
            get { return this.yomiGivenName; }
        }

        /// <summary>
        /// Gets the Yomi surname (last name) of the contact.
        /// </summary>
        public string YomiSurname
        {
            get { return this.yomiSurname; }
        }

        #endregion

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.Title:
                    this.title = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FirstName:
                    this.givenName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.MiddleName:
                    this.middleName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.LastName:
                    this.surname = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Suffix:
                    this.suffix = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Initials:
                    this.initials = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FullName:
                    this.fullName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.NickName:
                    this.nickname = reader.ReadElementValue();
                    return true;
                case XmlElementNames.YomiFirstName:
                    this.yomiGivenName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.YomiLastName:
                    this.yomiSurname = reader.ReadElementValue();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service"></param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.Title:
                        this.title = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.FirstName:
                        this.givenName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.MiddleName:
                        this.middleName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.LastName:
                        this.surname = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.Suffix:
                        this.suffix = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.Initials:
                        this.initials = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.FullName:
                        this.fullName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.NickName:
                        this.nickname = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.YomiFirstName:
                        this.yomiGivenName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.YomiLastName:
                        this.yomiSurname = jsonProperty.ReadAsString(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Title, this.Title);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FirstName, this.GivenName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MiddleName, this.MiddleName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LastName, this.Surname);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Suffix, this.Suffix);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Initials, this.Initials);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FullName, this.FullName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.NickName, this.NickName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.YomiFirstName, this.YomiGivenName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.YomiLastName, this.YomiSurname);
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service"></param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlElementNames.Title, this.Title);
            jsonProperty.Add(XmlElementNames.FirstName, this.GivenName);
            jsonProperty.Add(XmlElementNames.MiddleName, this.MiddleName);
            jsonProperty.Add(XmlElementNames.LastName, this.Surname);
            jsonProperty.Add(XmlElementNames.Suffix, this.Suffix);
            jsonProperty.Add(XmlElementNames.Initials, this.Initials);
            jsonProperty.Add(XmlElementNames.FullName, this.FullName);
            jsonProperty.Add(XmlElementNames.NickName, this.NickName);
            jsonProperty.Add(XmlElementNames.YomiFirstName, this.YomiGivenName);
            jsonProperty.Add(XmlElementNames.YomiLastName, this.YomiSurname);

            return jsonProperty;
        }
    }
}