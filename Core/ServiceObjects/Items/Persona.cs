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

    /// <summary>
    /// Represents a Persona. Properties available on Personas are defined in the PersonaSchema class.
    /// </summary>
    [Attachable]
    [ServiceObjectDefinition(XmlElementNames.Persona)]
    public class Persona : Item
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="Persona"/>. To bind to an existing Persona, use Persona.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the Persona will be bound.</param>
        public Persona(ExchangeService service)
            : base(service)
        {
            this.PersonaType = string.Empty;
            this.CreationTime = null;
            this.DisplayNameFirstLastHeader = string.Empty;
            this.DisplayNameLastFirstHeader = string.Empty;
            this.DisplayName = string.Empty;
            this.DisplayNameFirstLast = string.Empty;
            this.DisplayNameLastFirst = string.Empty;
            this.FileAs = string.Empty;
            this.Generation = string.Empty;
            this.DisplayNamePrefix = string.Empty;
            this.GivenName = string.Empty;
            this.Surname = string.Empty;
            this.Title = string.Empty;
            this.CompanyName = string.Empty;
            this.ImAddress = string.Empty;
            this.HomeCity = string.Empty;
            this.WorkCity = string.Empty;
            this.Alias = string.Empty;
            this.RelevanceScore = 0;

            // Remaining properties are initialized when the property definition is created in
            // PersonaSchema.cs.
        }

        /// <summary>
        /// Binds to an existing Persona and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the Persona.</param>
        /// <param name="id">The Id of the Persona to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A Persona instance representing the Persona corresponding to the specified Id.</returns>
        public static new Persona Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet)
        {
            return service.BindToItem<Persona>(id, propertySet);
        }

        /// <summary>
        /// Binds to an existing Persona and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the Persona.</param>
        /// <param name="id">The Id of the Persona to bind to.</param>
        /// <returns>A Persona instance representing the Persona corresponding to the specified Id.</returns>
        public static new Persona Bind(ExchangeService service, ItemId id)
        {
            return Persona.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return PersonaSchema.Instance;
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013_SP1;
        }

        /// <summary>
        /// The property definition for the Id of this object.
        /// </summary>
        /// <returns>A PropertyDefinition instance.</returns>
        internal override PropertyDefinition GetIdPropertyDefinition()
        {
            return PersonaSchema.PersonaId;
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
        }

        #region Properties

        /// <summary>
        /// Gets the persona id
        /// </summary>
        public ItemId PersonaId
        {
            get { return (ItemId)this.PropertyBag[this.GetIdPropertyDefinition()]; }
            set { this.PropertyBag[this.GetIdPropertyDefinition()] = value; }
        }

        /// <summary>
        /// Gets the persona type
        /// </summary>
        public string PersonaType
        {
            get { return (string)this.PropertyBag[PersonaSchema.PersonaType]; }
            set { this.PropertyBag[PersonaSchema.PersonaType] = value; }
        }

        /// <summary>
        /// Gets the creation time of the underlying contact
        /// </summary>
        public DateTime? CreationTime
        {
            get { return (DateTime?)this.PropertyBag[PersonaSchema.CreationTime]; }
            set { this.PropertyBag[PersonaSchema.CreationTime] = value; }
        }

        /// <summary>
        /// Gets the header of the FirstLast display name
        /// </summary>
        public string DisplayNameFirstLastHeader
        {
            get { return (string)this.PropertyBag[PersonaSchema.DisplayNameFirstLastHeader]; }
            set { this.PropertyBag[PersonaSchema.DisplayNameFirstLastHeader] = value; }
        }

        /// <summary>
        /// Gets the header of the LastFirst display name
        /// </summary>
        public string DisplayNameLastFirstHeader
        {
            get { return (string)this.PropertyBag[PersonaSchema.DisplayNameLastFirstHeader]; }
            set { this.PropertyBag[PersonaSchema.DisplayNameLastFirstHeader] = value; }
        }

        /// <summary>
        /// Gets the display name
        /// </summary>
        public string DisplayName
        {
            get { return (string)this.PropertyBag[PersonaSchema.DisplayName]; }
            set { this.PropertyBag[PersonaSchema.DisplayName] = value; }
        }

        /// <summary>
        /// Gets the display name in first last order
        /// </summary>
        public string DisplayNameFirstLast
        {
            get { return (string)this.PropertyBag[PersonaSchema.DisplayNameFirstLast]; }
            set { this.PropertyBag[PersonaSchema.DisplayNameFirstLast] = value; }
        }

        /// <summary>
        /// Gets the display name in last first order
        /// </summary>
        public string DisplayNameLastFirst
        {
            get { return (string)this.PropertyBag[PersonaSchema.DisplayNameLastFirst]; }
            set { this.PropertyBag[PersonaSchema.DisplayNameLastFirst] = value; }
        }

        /// <summary>
        /// Gets the name under which this Persona is filed as. FileAs can be manually set or
        /// can be automatically calculated based on the value of the FileAsMapping property.
        /// </summary>
        public string FileAs
        {
            get { return (string)this.PropertyBag[PersonaSchema.FileAs]; }
            set { this.PropertyBag[PersonaSchema.FileAs] = value; }
        }

        /// <summary>
        /// Gets the generation of the Persona
        /// </summary>
        public string Generation
        {
            get { return (string)this.PropertyBag[PersonaSchema.Generation]; }
            set { this.PropertyBag[PersonaSchema.Generation] = value; }
        }

        /// <summary>
        /// Gets the DisplayNamePrefix of the Persona
        /// </summary>
        public string DisplayNamePrefix
        {
            get { return (string)this.PropertyBag[PersonaSchema.DisplayNamePrefix]; }
            set { this.PropertyBag[PersonaSchema.DisplayNamePrefix] = value; }
        }

        /// <summary>
        /// Gets the given name of the Persona
        /// </summary>
        public string GivenName
        {
            get { return (string)this.PropertyBag[PersonaSchema.GivenName]; }
            set { this.PropertyBag[PersonaSchema.GivenName] = value; }
        }

        /// <summary>
        /// Gets the surname of the Persona
        /// </summary>
        public string Surname
        {
            get { return (string)this.PropertyBag[PersonaSchema.Surname]; }
            set { this.PropertyBag[PersonaSchema.Surname] = value; }
        }

        /// <summary>
        /// Gets the Persona's title
        /// </summary>
        public string Title
        {
            get { return (string)this.PropertyBag[PersonaSchema.Title]; }
            set { this.PropertyBag[PersonaSchema.Title] = value; }
        }

        /// <summary>
        /// Gets the company name of the Persona
        /// </summary>
        public string CompanyName
        {
            get { return (string)this.PropertyBag[PersonaSchema.CompanyName]; }
            set { this.PropertyBag[PersonaSchema.CompanyName] = value; }
        }

        /// <summary>
        /// Gets the email of the persona
        /// </summary>
        public PersonaEmailAddress EmailAddress
        {
            get { return (PersonaEmailAddress)this.PropertyBag[PersonaSchema.EmailAddress]; }
            set { this.PropertyBag[PersonaSchema.EmailAddress] = value; }
        }

        /// <summary>
        /// Gets the list of e-mail addresses of the contact
        /// </summary>
        public PersonaEmailAddressCollection EmailAddresses
        {
            get { return (PersonaEmailAddressCollection)this.PropertyBag[PersonaSchema.EmailAddresses]; }
            set { this.PropertyBag[PersonaSchema.EmailAddresses] = value; }
        }

        /// <summary>
        /// Gets the IM address of the persona
        /// </summary>
        public string ImAddress
        {
            get { return (string)this.PropertyBag[PersonaSchema.ImAddress]; }
            set { this.PropertyBag[PersonaSchema.ImAddress] = value; }
        }

        /// <summary>
        /// Gets the city of the Persona's home
        /// </summary>
        public string HomeCity
        {
            get { return (string)this.PropertyBag[PersonaSchema.HomeCity]; }
            set { this.PropertyBag[PersonaSchema.HomeCity] = value; }
        }

        /// <summary>
        /// Gets the city of the Persona's work place
        /// </summary>
        public string WorkCity
        {
            get { return (string)this.PropertyBag[PersonaSchema.WorkCity]; }
            set { this.PropertyBag[PersonaSchema.WorkCity] = value; }
        }

        /// <summary>
        /// Gets the alias of the Persona
        /// </summary>
        public string Alias
        {
            get { return (string)this.PropertyBag[PersonaSchema.Alias]; }
            set { this.PropertyBag[PersonaSchema.Alias] = value; }
        }

        /// <summary>
        /// Gets the relevance score
        /// </summary>
        public int RelevanceScore
        {
            get { return (int)this.PropertyBag[PersonaSchema.RelevanceScore]; }
            set { this.PropertyBag[PersonaSchema.RelevanceScore] = value; }
        }

        /// <summary>
        /// Gets the list of attributions
        /// </summary>
        public AttributionCollection Attributions
        {
            get { return (AttributionCollection)this.PropertyBag[PersonaSchema.Attributions]; }
            set { this.PropertyBag[PersonaSchema.Attributions] = value; }
        }

        /// <summary>
        /// Gets the list of office locations
        /// </summary>
        public AttributedStringCollection OfficeLocations
        {
            get { return (AttributedStringCollection)this.PropertyBag[PersonaSchema.OfficeLocations]; }
            set { this.PropertyBag[PersonaSchema.OfficeLocations] = value; }
        }

        /// <summary>
        /// Gets the list of IM addresses of the persona
        /// </summary>
        public AttributedStringCollection ImAddresses
        {
            get { return (AttributedStringCollection)this.PropertyBag[PersonaSchema.ImAddresses]; }
            set { this.PropertyBag[PersonaSchema.ImAddresses] = value; }
        }

        /// <summary>
        /// Gets the list of departments of the persona
        /// </summary>
        public AttributedStringCollection Departments
        {
            get { return (AttributedStringCollection)this.PropertyBag[PersonaSchema.Departments]; }
            set { this.PropertyBag[PersonaSchema.Departments] = value; }
        }

        /// <summary>
        /// Gets the list of photo URLs
        /// </summary>
        public AttributedStringCollection ThirdPartyPhotoUrls
        {
            get { return (AttributedStringCollection)this.PropertyBag[PersonaSchema.ThirdPartyPhotoUrls]; }
            set { this.PropertyBag[PersonaSchema.ThirdPartyPhotoUrls] = value; }
        }

        #endregion
    }
}