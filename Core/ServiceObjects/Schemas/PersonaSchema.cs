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
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Persona schema
    /// </summary>
    [Schema]
    public class PersonaSchema : ItemSchema
    {
        /// <summary>
        /// FieldURIs for persona.
        /// </summary>
        private static class FieldUris
        {
            public const string PersonaId = "persona:PersonaId";
            public const string PersonaType = "persona:PersonaType";
            public const string CreationTime = "persona:CreationTime";
            public const string DisplayNameFirstLastHeader = "persona:DisplayNameFirstLastHeader";
            public const string DisplayNameLastFirstHeader = "persona:DisplayNameLastFirstHeader";
            public const string DisplayName = "persona:DisplayName";
            public const string DisplayNameFirstLast = "persona:DisplayNameFirstLast";
            public const string DisplayNameLastFirst = "persona:DisplayNameLastFirst";
            public const string FileAs = "persona:FileAs";
            public const string Generation = "persona:Generation";
            public const string DisplayNamePrefix = "persona:DisplayNamePrefix";
            public const string GivenName = "persona:GivenName";
            public const string Surname = "persona:Surname";
            public const string Title = "persona:Title";
            public const string CompanyName = "persona:CompanyName";
            public const string EmailAddress = "persona:EmailAddress";
            public const string EmailAddresses = "persona:EmailAddresses";
            public const string ImAddress = "persona:ImAddress";
            public const string HomeCity = "persona:HomeCity";
            public const string WorkCity = "persona:WorkCity";
            public const string Alias = "persona:Alias";
            public const string RelevanceScore = "persona:RelevanceScore";
            public const string Attributions = "persona:Attributions";
            public const string OfficeLocations = "persona:OfficeLocations";
            public const string ImAddresses = "persona:ImAddresses";
            public const string Departments = "persona:Departments";
            public const string ThirdPartyPhotoUrls = "persona:ThirdPartyPhotoUrls";
        }

        /// <summary>
        /// Defines the PersonaId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PersonaId =
            new ComplexPropertyDefinition<ItemId>(
                XmlElementNames.PersonaId,
                FieldUris.PersonaId,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new ItemId(); });

        /// <summary>
        /// Defines the PersonaType property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PersonaType =
            new StringPropertyDefinition(
                XmlElementNames.PersonaType,
                FieldUris.PersonaType,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the CreationTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition CreationTime =
            new DateTimePropertyDefinition(
                XmlElementNames.CreationTime,
                FieldUris.CreationTime,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the DisplayNameFirstLastHeader property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayNameFirstLastHeader =
            new StringPropertyDefinition(
                XmlElementNames.DisplayNameFirstLastHeader,
                FieldUris.DisplayNameFirstLastHeader,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the DisplayNameLastFirstHeader property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayNameLastFirstHeader =
            new StringPropertyDefinition(
                XmlElementNames.DisplayNameLastFirstHeader,
                FieldUris.DisplayNameLastFirstHeader,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the DisplayName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayName =
            new StringPropertyDefinition(
                XmlElementNames.DisplayName,
                FieldUris.DisplayName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the DisplayNameFirstLast property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayNameFirstLast =
            new StringPropertyDefinition(
                XmlElementNames.DisplayNameFirstLast,
                FieldUris.DisplayNameFirstLast,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the DisplayNameLastFirst property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayNameLastFirst =
            new StringPropertyDefinition(
                XmlElementNames.DisplayNameLastFirst,
                FieldUris.DisplayNameLastFirst,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the FileAs property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FileAs =
            new StringPropertyDefinition(
                XmlElementNames.FileAs,
                FieldUris.FileAs,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the Generation property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Generation =
            new StringPropertyDefinition(
                XmlElementNames.Generation,
                FieldUris.Generation,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the DisplayNamePrefix property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayNamePrefix =
            new StringPropertyDefinition(
                XmlElementNames.DisplayNamePrefix,
                FieldUris.DisplayNamePrefix,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the GivenName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GivenName =
            new StringPropertyDefinition(
                XmlElementNames.GivenName,
                FieldUris.GivenName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the Surname property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Surname =
            new StringPropertyDefinition(
                XmlElementNames.Surname,
                FieldUris.Surname,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the Title property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Title =
            new StringPropertyDefinition(
                XmlElementNames.Title,
                FieldUris.Title,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the CompanyName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition CompanyName =
            new StringPropertyDefinition(
                XmlElementNames.CompanyName,
                FieldUris.CompanyName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the EmailAddress property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EmailAddress =
            new ComplexPropertyDefinition<PersonaEmailAddress>(
                XmlElementNames.EmailAddress,
                FieldUris.EmailAddress,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new PersonaEmailAddress(); });

        /// <summary>
        /// Defines the EmailAddresses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EmailAddresses =
            new ComplexPropertyDefinition<PersonaEmailAddressCollection>(
                XmlElementNames.EmailAddresses,
                FieldUris.EmailAddresses,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new PersonaEmailAddressCollection(); });

        /// <summary>
        /// Defines the ImAddress property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ImAddress =
            new StringPropertyDefinition(
                XmlElementNames.ImAddress,
                FieldUris.ImAddress,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the HomeCity property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition HomeCity =
            new StringPropertyDefinition(
                XmlElementNames.HomeCity,
                FieldUris.HomeCity,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the WorkCity property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition WorkCity =
            new StringPropertyDefinition(
                XmlElementNames.WorkCity,
                FieldUris.WorkCity,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the Alias property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Alias =
            new StringPropertyDefinition(
                XmlElementNames.Alias,
                FieldUris.Alias,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1);

        /// <summary>
        /// Defines the RelevanceScore property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition RelevanceScore =
            new IntPropertyDefinition(
                XmlElementNames.RelevanceScore,
                FieldUris.RelevanceScore,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                true);

        /// <summary>
        /// Defines the Attributions property
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Attributions =
            new ComplexPropertyDefinition<AttributionCollection>(
                XmlElementNames.Attributions,
                FieldUris.Attributions,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new AttributionCollection(); });

        /// <summary>
        /// Defines the OfficeLocations property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OfficeLocations =
            new ComplexPropertyDefinition<AttributedStringCollection>(
                XmlElementNames.OfficeLocations,
                FieldUris.OfficeLocations,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new AttributedStringCollection(); });

        /// <summary>
        /// Defines the ImAddresses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ImAddresses =
            new ComplexPropertyDefinition<AttributedStringCollection>(
                XmlElementNames.ImAddresses,
                FieldUris.ImAddresses,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new AttributedStringCollection(); });

        /// <summary>
        /// Defines the Departments property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Departments =
            new ComplexPropertyDefinition<AttributedStringCollection>(
                XmlElementNames.Departments,
                FieldUris.Departments,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new AttributedStringCollection(); });

        /// <summary>
        /// Defines the ThirdPartyPhotoUrls property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ThirdPartyPhotoUrls =
            new ComplexPropertyDefinition<AttributedStringCollection>(
                XmlElementNames.ThirdPartyPhotoUrls,
                FieldUris.ThirdPartyPhotoUrls,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013_SP1,
                delegate() { return new AttributedStringCollection(); });

        // This must be declared after the property definitions
        internal static new readonly PersonaSchema Instance = new PersonaSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(PersonaId);
            this.RegisterProperty(PersonaType);
            this.RegisterProperty(CreationTime);
            this.RegisterProperty(DisplayNameFirstLastHeader);
            this.RegisterProperty(DisplayNameLastFirstHeader);
            this.RegisterProperty(DisplayName);
            this.RegisterProperty(DisplayNameFirstLast);
            this.RegisterProperty(DisplayNameLastFirst);
            this.RegisterProperty(FileAs);
            this.RegisterProperty(Generation);
            this.RegisterProperty(DisplayNamePrefix);
            this.RegisterProperty(GivenName);
            this.RegisterProperty(Surname);
            this.RegisterProperty(Title);
            this.RegisterProperty(CompanyName);
            this.RegisterProperty(EmailAddress);
            this.RegisterProperty(EmailAddresses);
            this.RegisterProperty(ImAddress);
            this.RegisterProperty(HomeCity);
            this.RegisterProperty(WorkCity);
            this.RegisterProperty(Alias);
            this.RegisterProperty(RelevanceScore);
            this.RegisterProperty(Attributions);
            this.RegisterProperty(OfficeLocations);
            this.RegisterProperty(ImAddresses);
            this.RegisterProperty(Departments);
            this.RegisterProperty(ThirdPartyPhotoUrls);
        }

        /// <summary>
        /// internal constructor
        /// </summary>
        internal PersonaSchema()
            : base()
        {
        }
    }
}