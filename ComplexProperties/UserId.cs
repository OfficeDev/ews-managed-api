#region License

// Exchange Web Services Managed API
// 
// Copyright (c) Microsoft Corporation
// All rights reserved. 
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this
// software and associated documentation files (the "Software"), to deal in the Software
// without restriction, including without limitation the rights to use, copy, modify, merge,
// publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
// to whom the Software is furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or
// substantial portions of the Software.

// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
// PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
// FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.

#endregion

//-----------------------------------------------------------------------
// <summary>Defines the UserId class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;

    /// <summary>
    /// Represents the Id of a user.
    /// </summary>
    public sealed class UserId : ComplexProperty
    {
        private string sID;
        private string primarySmtpAddress;
        private string displayName;
        private StandardUser? standardUser;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserId"/> class.
        /// </summary>
        public UserId()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserId"/> class.
        /// </summary>
        /// <param name="primarySmtpAddress">The primary SMTP address used to initialize the UserId.</param>
        public UserId(string primarySmtpAddress)
            : this()
        {
            this.primarySmtpAddress = primarySmtpAddress;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserId"/> class.
        /// </summary>
        /// <param name="standardUser">The StandardUser value used to initialize the UserId.</param>
        public UserId(StandardUser standardUser)
            : this()
        {
            this.standardUser = standardUser;
        }

        /// <summary>
        /// Determines whether this instance is valid.
        /// </summary>
        /// <returns><c>true</c> if this instance is valid; otherwise, <c>false</c>.</returns>
        internal bool IsValid()
        {
            return this.StandardUser.HasValue || !string.IsNullOrEmpty(this.PrimarySmtpAddress) || !string.IsNullOrEmpty(this.SID);
        }

        /// <summary>
        /// Gets or sets the SID of the user.
        /// </summary>
        public string SID
        {
            get { return this.sID; }
            set { this.SetFieldValue<string>(ref this.sID, value); }
        }

        /// <summary>
        /// Gets or sets the primary SMTP address or the user.
        /// </summary>
        public string PrimarySmtpAddress
        {
            get { return this.primarySmtpAddress; }
            set { this.SetFieldValue<string>(ref this.primarySmtpAddress, value); }
        }

        /// <summary>
        /// Gets or sets the display name of the user.
        /// </summary>
        public string DisplayName
        {
            get { return this.displayName; }
            set { this.SetFieldValue<string>(ref this.displayName, value); }
        }

        /// <summary>
        /// Gets or sets a value indicating which standard user the user represents.
        /// </summary>
        public StandardUser? StandardUser
        {
            get { return this.standardUser; }
            set { this.SetFieldValue<StandardUser?>(ref this.standardUser, value); }
        }

        /// <summary>
        /// Implements an implicit conversion between a string representing a primary SMTP address and UserId.
        /// </summary>
        /// <param name="primarySmtpAddress">The string representing a primary SMTP address.</param>
        /// <returns>A UserId initialized with the specified primary SMTP address.</returns>
        public static implicit operator UserId(string primarySmtpAddress)
        {
            return new UserId(primarySmtpAddress);
        }

        /// <summary>
        /// Implements an implicit conversion between StandardUser and UserId.
        /// </summary>
        /// <param name="standardUser">The standard user used to initialize the user Id.</param>
        /// <returns>A UserId initialized with the specified standard user value.</returns>
        public static implicit operator UserId(StandardUser standardUser)
        {
            return new UserId(standardUser);
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.SID:
                    this.sID = reader.ReadValue();
                    return true;
                case XmlElementNames.PrimarySmtpAddress:
                    this.primarySmtpAddress = reader.ReadValue();
                    return true;
                case XmlElementNames.DisplayName:
                    this.displayName = reader.ReadValue();
                    return true;
                case XmlElementNames.DistinguishedUser:
                    this.standardUser = reader.ReadValue<StandardUser>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            foreach (string key in jsonProperty.Keys)
            {
                switch (key)
                {
                    case XmlElementNames.SID:
                        this.sID = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.PrimarySmtpAddress:
                        this.primarySmtpAddress = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.DisplayName:
                        this.displayName = jsonProperty.ReadAsString(key);
                        break;
                    case XmlElementNames.DistinguishedUser:
                        this.standardUser = jsonProperty.ReadEnumValue<StandardUser>(key);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.SID,
                this.SID);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.PrimarySmtpAddress,
                this.PrimarySmtpAddress);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.DisplayName,
                this.DisplayName);

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.DistinguishedUser,
                this.StandardUser);
        }

        /// <summary>
        /// Serializes the property to a Json value.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal override object InternalToJson(ExchangeService service)
        {
            JsonObject jsonProperty = new JsonObject();

            jsonProperty.Add(XmlElementNames.SID, this.SID);
            jsonProperty.Add(XmlElementNames.PrimarySmtpAddress, this.PrimarySmtpAddress);
            jsonProperty.Add(XmlElementNames.DisplayName, this.DisplayName);

            if (this.StandardUser.HasValue)
            {
                jsonProperty.Add(XmlElementNames.DistinguishedUser, this.StandardUser.Value);
            }

            return jsonProperty;
        }
    }
}
