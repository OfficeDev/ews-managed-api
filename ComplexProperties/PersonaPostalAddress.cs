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
    using System.Xml;

    /// <summary>
    /// Represents PersonaPostalAddress.
    /// </summary>
    public sealed class PersonaPostalAddress : ComplexProperty
    {
        private string street;
        private string city;
        private string state;
        private string country;
        private string postalCode;
        private string postOfficeBox;
        private string type;
        private double? latitude;
        private double? longitude;
        private double? accuracy;
        private double? altitude;
        private double? altitudeAccuracy;
        private string formattedAddress;
        private string uri;
        private LocationSource source;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonaPostalAddress"/> class.
        /// </summary>
        internal PersonaPostalAddress()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonaPostalAddress"/> class.
        /// </summary>
        /// <param name="street">The Street Address.</param>
        /// <param name="city">The City value.</param>
        /// <param name="state">The State value.</param>
        /// <param name="country">The country value.</param>
        /// <param name="postalCode">The postal code value.</param>
        /// <param name="postOfficeBox">The Post Office Box.</param>
        /// <param name="locationSource">The location Source.</param>
        /// <param name="locationUri">The location Uri.</param>
        /// <param name="formattedAddress">The location street Address in formatted address.</param>
        /// <param name="latitude">The location latitude.</param>
        /// <param name="longitude">The location longitude.</param>
        /// <param name="accuracy">The location accuracy.</param>
        /// <param name="altitude">The location altitude.</param>
        /// <param name="altitudeAccuracy">The location altitude Accuracy.</param>
        public PersonaPostalAddress(
            string street, 
            string city, 
            string state, 
            string country,
            string postalCode, 
            string postOfficeBox, 
            LocationSource locationSource, 
            string locationUri, 
            string formattedAddress, 
            double latitude, 
            double longitude, 
            double accuracy, 
            double altitude, 
            double altitudeAccuracy)
            : this()
        {
            this.street = street;
            this.city = city;
            this.state = state;
            this.country = country;
            this.postalCode = postalCode;
            this.postOfficeBox = postOfficeBox;
            this.latitude = latitude;
            this.longitude = longitude;
            this.source = locationSource;
            this.uri = locationUri;
            this.formattedAddress = formattedAddress;
            this.accuracy = accuracy;
            this.altitude = altitude;
            this.altitudeAccuracy = altitudeAccuracy;
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
                case XmlElementNames.Street:
                    this.street = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.City:
                    this.city = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.State:
                    this.state = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.Country:
                    this.country = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PostalCode:
                    this.postalCode = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PostOfficeBox:
                    this.postOfficeBox = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PostalAddressType:
                    this.type = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.Latitude:
                    this.latitude = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.Longitude:
                    this.longitude = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.Accuracy:
                    this.accuracy = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.Altitude:
                    this.altitude = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.AltitudeAccuracy:
                    this.altitudeAccuracy = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.FormattedAddress:
                    this.formattedAddress = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.LocationUri:
                    this.uri = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.LocationSource:
                    this.source = reader.ReadValue<LocationSource>();
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
        {
            do
            {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                {
                    this.TryReadElementFromXml(reader);
                }
            }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.PersonaPostalAddress));
        }

        /// <summary>
        /// Gets or sets the street.
        /// </summary>
        public string Street
        {
            get { return this.street; }
            set { this.SetFieldValue<string>(ref this.street, value); }
        }
        
        /// <summary>
        /// Gets or sets the City.
        /// </summary>
        public string City
        {
            get { return this.city; }
            set { this.SetFieldValue<string>(ref this.city, value); }
        }
        
        /// <summary>
        /// Gets or sets the state.
        /// </summary>
        public string State
        {
            get { return this.state; }
            set { this.SetFieldValue<string>(ref this.state, value); }
        }

        /// <summary>
        /// Gets or sets the Country.
        /// </summary>
        public string Country
        {
            get { return this.country; }
            set { this.SetFieldValue<string>(ref this.country, value); }
        }
        
        /// <summary>
        /// Gets or sets the postalCode.
        /// </summary>
        public string PostalCode
        {
            get { return this.postalCode; }
            set { this.SetFieldValue<string>(ref this.postalCode, value); }
        }   
 
        /// <summary>
        /// Gets or sets the postOfficeBox.
        /// </summary>
        public string PostOfficeBox
        {
            get { return this.postOfficeBox; }
            set { this.SetFieldValue<string>(ref this.postOfficeBox, value); }
        }       
        
        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        public string Type
        {
            get { return this.type; }
            set { this.SetFieldValue<string>(ref this.type, value); }
        }    

        /// <summary>
        /// Gets or sets the location source type.
        /// </summary>
        public LocationSource Source
        {
            get { return this.source; }
            set { this.SetFieldValue<LocationSource>(ref this.source, value); }
        }

        /// <summary>
        /// Gets or sets the location Uri.
        /// </summary>
        public string Uri
        {
            get { return this.uri; }
            set { this.SetFieldValue<string>(ref this.uri, value); }
        }

        /// <summary>
        /// Gets or sets a value indicating location latitude.
        /// </summary>
        public double? Latitude
        {
            get { return this.latitude; }
            set { this.SetFieldValue<double?>(ref this.latitude, value); }
        }

        /// <summary>
        /// Gets or sets a value indicating location longitude.
        /// </summary>
        public double? Longitude
        {
            get { return this.longitude; }
            set { this.SetFieldValue<double?>(ref this.longitude, value); }
        }

        /// <summary>
        /// Gets or sets the location accuracy.
        /// </summary>
        public double? Accuracy
        {
            get { return this.accuracy; }
            set { this.SetFieldValue<double?>(ref this.accuracy, value); }
        }

        /// <summary>
        /// Gets or sets the location altitude.
        /// </summary>
        public double? Altitude
        {
            get { return this.altitude; }
            set { this.SetFieldValue<double?>(ref this.altitude, value); }
        }

        /// <summary>
        /// Gets or sets the location altitude accuracy.
        /// </summary>
        public double? AltitudeAccuracy
        {
            get { return this.altitudeAccuracy; }
            set { this.SetFieldValue<double?>(ref this.altitudeAccuracy, value); }
        }

        /// <summary>
        /// Gets or sets the street address.
        /// </summary>
        public string FormattedAddress
        {
            get { return this.formattedAddress; }
            set { this.SetFieldValue<string>(ref this.formattedAddress, value); }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Street, this.street);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.City, this.city);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.State, this.state);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Country, this.country);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostalCode, this.postalCode);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostOfficeBox, this.postOfficeBox);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostalAddressType, this.type);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Latitude, this.latitude);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Longitude, this.longitude);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Accuracy, this.accuracy);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Altitude, this.altitude);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AltitudeAccuracy, this.altitudeAccuracy);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FormattedAddress, this.formattedAddress);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationUri, this.uri);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationSource, this.source);
        }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.PersonaPostalAddress);

            this.WriteElementsToXml(writer);

            writer.WriteEndElement(); // xmlElementName
        }
    }
}