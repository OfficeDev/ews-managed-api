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
    /// <summary>
    /// Represents a mailbox reference.
    /// </summary>
    public class Mailbox : ComplexProperty, ISearchStringProvider
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Mailbox"/> class.
        /// </summary>
        public Mailbox()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Mailbox"/> class.
        /// </summary>
        /// <param name="smtpAddress">The primary SMTP address of the mailbox.</param>
        public Mailbox(string smtpAddress)
            : this()
        {
            this.Address = smtpAddress;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Mailbox"/> class.
        /// </summary>
        /// <param name="address">The address used to reference the user mailbox.</param>
        /// <param name="routingType">The routing type of the address used to reference the user mailbox.</param>
        public Mailbox(string address, string routingType)
            : this(address)
        {
            this.RoutingType = routingType;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// True if this instance is valid, false otherthise.
        /// </summary>
        /// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
        public bool IsValid
        {
            get { return !string.IsNullOrEmpty(this.Address); }
        }

        /// <summary>
        /// Gets or sets the address used to refer to the user mailbox.
        /// </summary>
        public string Address
        {
            get; set;
        }

        /// <summary>
        /// Gets or sets the routing type of the address used to refer to the user mailbox.
        /// </summary>
        public string RoutingType
        {
            get; set;
        }

        #endregion

        #region Operator overloads
        
        /// <summary>
        /// Defines an implicit conversion between a string representing an SMTP address and Mailbox.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address to convert to EmailAddress.</param>
        /// <returns>A Mailbox initialized with the specified SMTP address.</returns>
        public static implicit operator Mailbox(string smtpAddress)
        {
            return new Mailbox(smtpAddress);
        }

        #endregion

        #region Xml Methods

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.EmailAddress:
                    this.Address = reader.ReadElementValue();
                    return true;
                case XmlElementNames.RoutingType:
                    this.RoutingType = reader.ReadElementValue();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.EmailAddress, this.Address);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.RoutingType, this.RoutingType);
        }

        #endregion

        #region Json Methods

        /// <summary>
        /// Loads from json.
        /// </summary>
        /// <param name="jsonProperty">The json property.</param>
        /// <param name="service">The service.</param>
        internal override void LoadFromJson(JsonObject jsonProperty, ExchangeService service)
        {
            if (jsonProperty.ContainsKey(XmlElementNames.EmailAddress))
            {
                this.Address = jsonProperty.ReadAsString(XmlElementNames.EmailAddress);
            }

            if (jsonProperty.ContainsKey(XmlElementNames.RoutingType))
            {
                this.RoutingType = jsonProperty.ReadAsString(XmlElementNames.RoutingType);
            }
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
            JsonObject jsonObject = new JsonObject();

            jsonObject.Add(XmlElementNames.EmailAddress, this.Address);
            jsonObject.Add(XmlElementNames.RoutingType, this.RoutingType);

            return jsonObject;
        }

        #endregion

        #region ISearchStringProvider methods
        /// <summary>
        /// Get a string representation for using this instance in a search filter.
        /// </summary>
        /// <returns>String representation of instance.</returns>
        string ISearchStringProvider.GetSearchString()
        {
            return this.Address;
        }
        #endregion

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
        {
            base.InternalValidate();

            EwsUtilities.ValidateNonBlankStringParamAllowNull(this.Address, "address");
            EwsUtilities.ValidateNonBlankStringParamAllowNull(this.RoutingType, "routingType");
        }

        #region Object method overrides
        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">The <paramref name="obj"/> parameter is null.</exception>
        public override bool Equals(object obj)
        {
            if (object.ReferenceEquals(this, obj))
            {
                return true;
            }
            else
            {
                Mailbox other = obj as Mailbox;

                if (other == null)
                {
                    return false;
                }
                else if (((this.Address == null) && (other.Address == null)) ||
                         ((this.Address != null) && this.Address.Equals(other.Address)))
                {
                    return ((this.RoutingType == null) && (other.RoutingType == null)) ||
                           ((this.RoutingType != null) && this.RoutingType.Equals(other.RoutingType));
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
        {
            if (!string.IsNullOrEmpty(this.Address))
            {
                int hashCode = this.Address.GetHashCode();

                if (!string.IsNullOrEmpty(this.RoutingType))
                {
                    hashCode ^= this.RoutingType.GetHashCode();
                }

                return hashCode;
            }
            else
            {
                return base.GetHashCode();
            }
        }

        /// <summary>
        /// Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.
        /// </returns>
        public override string ToString()
        {
            if (!this.IsValid)
            {
                return string.Empty;
            }
            else if (!string.IsNullOrEmpty(this.RoutingType))
            {
                return this.RoutingType + ":" + this.Address;
            }
            else
            {
                return this.Address;
            }
        }
        #endregion
    }
}