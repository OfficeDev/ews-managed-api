// ---------------------------------------------------------------------------
// <copyright file="ContactSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schem for contacts.
    /// </summary>
    [Schema]
    public class ContactSchema : ItemSchema
    {
        /// <summary>
        /// FieldURIs for contacts.
        /// </summary>
        private static class FieldUris
        {
            public const string FileAs = "contacts:FileAs";
            public const string FileAsMapping = "contacts:FileAsMapping";
            public const string DisplayName = "contacts:DisplayName";
            public const string GivenName = "contacts:GivenName";
            public const string Initials = "contacts:Initials";
            public const string MiddleName = "contacts:MiddleName";
            public const string NickName = "contacts:Nickname";
            public const string CompleteName = "contacts:CompleteName";
            public const string CompanyName = "contacts:CompanyName";
            public const string EmailAddress = "contacts:EmailAddress";
            public const string EmailAddresses = "contacts:EmailAddresses";
            public const string PhysicalAddresses = "contacts:PhysicalAddresses";
            public const string PhoneNumber = "contacts:PhoneNumber";
            public const string PhoneNumbers = "contacts:PhoneNumbers";
            public const string AssistantName = "contacts:AssistantName";
            public const string Birthday = "contacts:Birthday";
            public const string BusinessHomePage = "contacts:BusinessHomePage";
            public const string Children = "contacts:Children";
            public const string Companies = "contacts:Companies";
            public const string ContactSource = "contacts:ContactSource";
            public const string Department = "contacts:Department";
            public const string Generation = "contacts:Generation";
            public const string ImAddress = "contacts:ImAddress";
            public const string ImAddresses = "contacts:ImAddresses";
            public const string JobTitle = "contacts:JobTitle";
            public const string Manager = "contacts:Manager";
            public const string Mileage = "contacts:Mileage";
            public const string OfficeLocation = "contacts:OfficeLocation";
            public const string PhysicalAddressCity = "contacts:PhysicalAddress:City";
            public const string PhysicalAddressCountryOrRegion = "contacts:PhysicalAddress:CountryOrRegion";
            public const string PhysicalAddressState = "contacts:PhysicalAddress:State";
            public const string PhysicalAddressStreet = "contacts:PhysicalAddress:Street";
            public const string PhysicalAddressPostalCode = "contacts:PhysicalAddress:PostalCode";
            public const string PostalAddressIndex = "contacts:PostalAddressIndex";
            public const string Profession = "contacts:Profession";
            public const string SpouseName = "contacts:SpouseName";
            public const string Surname = "contacts:Surname";
            public const string WeddingAnniversary = "contacts:WeddingAnniversary";
            public const string HasPicture = "contacts:HasPicture";
            public const string PhoneticFullName = "contacts:PhoneticFullName";
            public const string PhoneticFirstName = "contacts:PhoneticFirstName";
            public const string PhoneticLastName = "contacts:PhoneticLastName";
            public const string Alias = "contacts:Alias";
            public const string Notes = "contacts:Notes";
            public const string Photo = "contacts:Photo";
            public const string UserSMIMECertificate = "contacts:UserSMIMECertificate";
            public const string MSExchangeCertificate = "contacts:MSExchangeCertificate";
            public const string DirectoryId = "contacts:DirectoryId";
            public const string ManagerMailbox = "contacts:ManagerMailbox";
            public const string DirectReports = "contacts:DirectReports";
        }

        /// <summary>
        /// Defines the FileAs property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FileAs =
            new StringPropertyDefinition(
                XmlElementNames.FileAs,
                FieldUris.FileAs,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the FileAsMapping property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FileAsMapping =
            new GenericPropertyDefinition<FileAsMapping>(
                XmlElementNames.FileAsMapping,
                FieldUris.FileAsMapping,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the DisplayName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayName =
            new StringPropertyDefinition(
                XmlElementNames.DisplayName,
                FieldUris.DisplayName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the GivenName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GivenName =
            new StringPropertyDefinition(
                XmlElementNames.GivenName,
                FieldUris.GivenName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Initials property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Initials =
            new StringPropertyDefinition(
                XmlElementNames.Initials,
                FieldUris.Initials,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the MiddleName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MiddleName =
            new StringPropertyDefinition(
                XmlElementNames.MiddleName,
                FieldUris.MiddleName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the NickName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition NickName =
            new StringPropertyDefinition(
                XmlElementNames.NickName,
                FieldUris.NickName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the CompleteName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition CompleteName =
            new ComplexPropertyDefinition<CompleteName>(
                XmlElementNames.CompleteName,
                FieldUris.CompleteName,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new CompleteName(); });

        /// <summary>
        /// Defines the CompanyName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition CompanyName =
            new StringPropertyDefinition(
                XmlElementNames.CompanyName,
                FieldUris.CompanyName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the EmailAddresses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EmailAddresses =
            new ComplexPropertyDefinition<EmailAddressDictionary>(
                XmlElementNames.EmailAddresses,
                FieldUris.EmailAddresses,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddressDictionary(); });

        /// <summary>
        /// Defines the PhysicalAddresses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PhysicalAddresses =
            new ComplexPropertyDefinition<PhysicalAddressDictionary>(
                XmlElementNames.PhysicalAddresses,
                FieldUris.PhysicalAddresses,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new PhysicalAddressDictionary(); });

        /// <summary>
        /// Defines the PhoneNumbers property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PhoneNumbers =
            new ComplexPropertyDefinition<PhoneNumberDictionary>(
                XmlElementNames.PhoneNumbers,
                FieldUris.PhoneNumbers,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new PhoneNumberDictionary(); });

        /// <summary>
        /// Defines the AssistantName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition AssistantName =
            new StringPropertyDefinition(
                XmlElementNames.AssistantName,
                FieldUris.AssistantName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Birthday property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Birthday =
            new DateTimePropertyDefinition(
                XmlElementNames.Birthday,
                FieldUris.Birthday,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the BusinessHomePage property.
        /// </summary>
        /// <remarks>
        /// Defined as anyURI in the EWS schema. String is fine here.
        /// </remarks>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition BusinessHomePage =
            new StringPropertyDefinition(
                XmlElementNames.BusinessHomePage,
                FieldUris.BusinessHomePage,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Children property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Children =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.Children,
                FieldUris.Children,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the Companies property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Companies =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.Companies,
                FieldUris.Companies,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the ContactSource property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ContactSource =
            new GenericPropertyDefinition<ContactSource>(
                XmlElementNames.ContactSource,
                FieldUris.ContactSource,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Department property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Department =
            new StringPropertyDefinition(
                XmlElementNames.Department,
                FieldUris.Department,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Generation property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Generation =
            new StringPropertyDefinition(
                XmlElementNames.Generation,
                FieldUris.Generation,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ImAddresses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ImAddresses =
            new ComplexPropertyDefinition<ImAddressDictionary>(
                XmlElementNames.ImAddresses,
                FieldUris.ImAddresses,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new ImAddressDictionary(); });

        /// <summary>
        /// Defines the JobTitle property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition JobTitle =
            new StringPropertyDefinition(
                XmlElementNames.JobTitle,
                FieldUris.JobTitle,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Manager property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Manager =
            new StringPropertyDefinition(
                XmlElementNames.Manager,
                FieldUris.Manager,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Mileage property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Mileage =
            new StringPropertyDefinition(
                XmlElementNames.Mileage,
                FieldUris.Mileage,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the OfficeLocation property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition OfficeLocation =
            new StringPropertyDefinition(
                XmlElementNames.OfficeLocation,
                FieldUris.OfficeLocation,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the PostalAddressIndex property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PostalAddressIndex =
            new GenericPropertyDefinition<PhysicalAddressIndex>(
                XmlElementNames.PostalAddressIndex,
                FieldUris.PostalAddressIndex,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Profession property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Profession =
            new StringPropertyDefinition(
                XmlElementNames.Profession,
                FieldUris.Profession,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the SpouseName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition SpouseName =
            new StringPropertyDefinition(
                XmlElementNames.SpouseName,
                FieldUris.SpouseName,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Surname property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Surname =
            new StringPropertyDefinition(
                XmlElementNames.Surname,
                FieldUris.Surname,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the WeddingAnniversary property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition WeddingAnniversary =
            new DateTimePropertyDefinition(
                XmlElementNames.WeddingAnniversary,
                FieldUris.WeddingAnniversary,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the HasPicture property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition HasPicture =
            new BoolPropertyDefinition(
                XmlElementNames.HasPicture,
                FieldUris.HasPicture,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010);
        
        #region Directory Only Properties

        /// <summary>
        /// Defines the PhoneticFullName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PhoneticFullName =
            new StringPropertyDefinition(
                XmlElementNames.PhoneticFullName,
                FieldUris.PhoneticFullName,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the PhoneticFirstName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PhoneticFirstName =
            new StringPropertyDefinition(
                XmlElementNames.PhoneticFirstName,
                FieldUris.PhoneticFirstName,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the PhoneticLastName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PhoneticLastName =
            new StringPropertyDefinition(
                XmlElementNames.PhoneticLastName,
                FieldUris.PhoneticLastName,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);
          
        /// <summary>
        /// Defines the Alias property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Alias =
            new StringPropertyDefinition(
                XmlElementNames.Alias,
                FieldUris.Alias,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the Notes property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Notes =
            new StringPropertyDefinition(
                XmlElementNames.Notes,
                FieldUris.Notes,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the Photo property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Photo =
            new ByteArrayPropertyDefinition(
                XmlElementNames.Photo, 
                FieldUris.Photo, 
                PropertyDefinitionFlags.CanFind, 
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the UserSMIMECertificate property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition UserSMIMECertificate =
            new ComplexPropertyDefinition<ByteArrayArray>(
                XmlElementNames.UserSMIMECertificate, 
                FieldUris.UserSMIMECertificate, 
                PropertyDefinitionFlags.CanFind, 
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new ByteArrayArray(); });

        /// <summary>
        /// Defines the MSExchangeCertificate property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MSExchangeCertificate =
            new ComplexPropertyDefinition<ByteArrayArray>(
                XmlElementNames.MSExchangeCertificate,
                FieldUris.MSExchangeCertificate, 
                PropertyDefinitionFlags.CanFind, 
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new ByteArrayArray(); });

        /// <summary>
        /// Defines the DirectoryId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DirectoryId =
            new StringPropertyDefinition(
                XmlElementNames.DirectoryId,
                FieldUris.DirectoryId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the ManagerMailbox property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ManagerMailbox =
            new ContainedPropertyDefinition<EmailAddress>(
                XmlElementNames.ManagerMailbox,
                FieldUris.ManagerMailbox,
                XmlElementNames.Mailbox,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new EmailAddress(); });

        /// <summary>
        /// Defines the DirectReports property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DirectReports =
            new ComplexPropertyDefinition<EmailAddressCollection>(
                XmlElementNames.DirectReports,
                FieldUris.DirectReports,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new EmailAddressCollection(); });

        #endregion

        #region Email addresses indexed properties

        /// <summary>
        /// Defines the EmailAddress1 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition EmailAddress1 =
            new IndexedPropertyDefinition(FieldUris.EmailAddress, "EmailAddress1");

        /// <summary>
        /// Defines the EmailAddress2 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition EmailAddress2 =
            new IndexedPropertyDefinition(FieldUris.EmailAddress, "EmailAddress2");

        /// <summary>
        /// Defines the EmailAddress3 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition EmailAddress3 =
            new IndexedPropertyDefinition(FieldUris.EmailAddress, "EmailAddress3");

        #endregion

        #region IM addresses indexed properties

        /// <summary>
        /// Defines the ImAddress1 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition ImAddress1 =
            new IndexedPropertyDefinition(FieldUris.ImAddress, "ImAddress1");

        /// <summary>
        /// Defines the ImAddress2 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition ImAddress2 =
            new IndexedPropertyDefinition(FieldUris.ImAddress, "ImAddress2");

        /// <summary>
        /// Defines the ImAddress3 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition ImAddress3 =
            new IndexedPropertyDefinition(FieldUris.ImAddress, "ImAddress3");

        #endregion

        #region Phone numbers indexed properties

        /// <summary>
        /// Defines the AssistentPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition AssistantPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "AssistantPhone");

        /// <summary>
        /// Defines the BusinessFax property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessFax =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "BusinessFax");

        /// <summary>
        /// Defines the BusinessPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "BusinessPhone");

        /// <summary>
        /// Defines the BusinessPhone2 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessPhone2 =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "BusinessPhone2");

        /// <summary>
        /// Defines the Callback property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition Callback =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "Callback");

        /// <summary>
        /// Defines the CarPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition CarPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "CarPhone");

        /// <summary>
        /// Defines the CompanyMainPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition CompanyMainPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "CompanyMainPhone");

        /// <summary>
        /// Defines the HomeFax property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomeFax =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "HomeFax");

        /// <summary>
        /// Defines the HomePhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomePhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "HomePhone");

        /// <summary>
        /// Defines the HomePhone2 property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomePhone2 =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "HomePhone2");

        /// <summary>
        /// Defines the Isdn property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition Isdn =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "Isdn");

        /// <summary>
        /// Defines the MobilePhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition MobilePhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "MobilePhone");

        /// <summary>
        /// Defines the OtherFax property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherFax =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "OtherFax");

        /// <summary>
        /// Defines the OtherTelephone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherTelephone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "OtherTelephone");

        /// <summary>
        /// Defines the Pager property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition Pager =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "Pager");

        /// <summary>
        /// Defines the PrimaryPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition PrimaryPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "PrimaryPhone");

        /// <summary>
        /// Defines the RadioPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition RadioPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "RadioPhone");

        /// <summary>
        /// Defines the Telex property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition Telex =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "Telex");

        /// <summary>
        /// Defines the TtyTddPhone property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition TtyTddPhone =
            new IndexedPropertyDefinition(FieldUris.PhoneNumber, "TtyTddPhone");

        #endregion

        #region Business address indexed properties

        /// <summary>
        /// Defines the BusinessAddressStreet property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessAddressStreet =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressStreet, "Business");

        /// <summary>
        /// Defines the BusinessAddressCity property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessAddressCity =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressCity, "Business");

        /// <summary>
        /// Defines the BusinessAddressState property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessAddressState =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressState, "Business");

        /// <summary>
        /// Defines the BusinessAddressCountryOrRegion property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessAddressCountryOrRegion =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressCountryOrRegion, "Business");

        /// <summary>
        /// Defines the BusinessAddressPostalCode property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition BusinessAddressPostalCode =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressPostalCode, "Business");

        #endregion

        #region Home address indexed properties

        /// <summary>
        /// Defines the HomeAddressStreet property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomeAddressStreet =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressStreet, "Home");

        /// <summary>
        /// Defines the HomeAddressCity property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomeAddressCity =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressCity, "Home");

        /// <summary>
        /// Defines the HomeAddressState property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomeAddressState =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressState, "Home");

        /// <summary>
        /// Defines the HomeAddressCountryOrRegion property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomeAddressCountryOrRegion =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressCountryOrRegion, "Home");

        /// <summary>
        /// Defines the HomeAddressPostalCode property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition HomeAddressPostalCode =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressPostalCode, "Home");

        #endregion

        #region Other address indexed properties

        /// <summary>
        /// Defines the OtherAddressStreet property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherAddressStreet =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressStreet, "Other");

        /// <summary>
        /// Defines the OtherAddressCity property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherAddressCity =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressCity, "Other");

        /// <summary>
        /// Defines the OtherAddressState property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherAddressState =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressState, "Other");

        /// <summary>
        /// Defines the OtherAddressCountryOrRegion property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherAddressCountryOrRegion =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressCountryOrRegion, "Other");

        /// <summary>
        /// Defines the OtherAddressPostalCode property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly IndexedPropertyDefinition OtherAddressPostalCode =
            new IndexedPropertyDefinition(FieldUris.PhysicalAddressPostalCode, "Other");

        #endregion

        // This must be declared after the property definitions
        internal static new readonly ContactSchema Instance = new ContactSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(FileAs);
            this.RegisterProperty(FileAsMapping);
            this.RegisterProperty(DisplayName);
            this.RegisterProperty(GivenName);
            this.RegisterProperty(Initials);
            this.RegisterProperty(MiddleName);
            this.RegisterProperty(NickName);
            this.RegisterProperty(CompleteName);
            this.RegisterProperty(CompanyName);
            this.RegisterProperty(EmailAddresses);
            this.RegisterProperty(PhysicalAddresses);
            this.RegisterProperty(PhoneNumbers);
            this.RegisterProperty(AssistantName);
            this.RegisterProperty(Birthday);
            this.RegisterProperty(BusinessHomePage);
            this.RegisterProperty(Children);
            this.RegisterProperty(Companies);
            this.RegisterProperty(ContactSource);
            this.RegisterProperty(Department);
            this.RegisterProperty(Generation);
            this.RegisterProperty(ImAddresses);
            this.RegisterProperty(JobTitle);
            this.RegisterProperty(Manager);
            this.RegisterProperty(Mileage);
            this.RegisterProperty(OfficeLocation);
            this.RegisterProperty(PostalAddressIndex);
            this.RegisterProperty(Profession);
            this.RegisterProperty(SpouseName);
            this.RegisterProperty(Surname);
            this.RegisterProperty(WeddingAnniversary);
            this.RegisterProperty(HasPicture);
            this.RegisterProperty(PhoneticFullName);
            this.RegisterProperty(PhoneticFirstName);
            this.RegisterProperty(PhoneticLastName);
            this.RegisterProperty(Alias);
            this.RegisterProperty(Notes);
            this.RegisterProperty(Photo);
            this.RegisterProperty(UserSMIMECertificate);
            this.RegisterProperty(MSExchangeCertificate);
            this.RegisterProperty(DirectoryId);
            this.RegisterProperty(ManagerMailbox);
            this.RegisterProperty(DirectReports);

            this.RegisterIndexedProperty(EmailAddress1);
            this.RegisterIndexedProperty(EmailAddress2);
            this.RegisterIndexedProperty(EmailAddress3);
            this.RegisterIndexedProperty(ImAddress1);
            this.RegisterIndexedProperty(ImAddress2);
            this.RegisterIndexedProperty(ImAddress3);
            this.RegisterIndexedProperty(AssistantPhone);
            this.RegisterIndexedProperty(BusinessFax);
            this.RegisterIndexedProperty(BusinessPhone);
            this.RegisterIndexedProperty(BusinessPhone2);
            this.RegisterIndexedProperty(Callback);
            this.RegisterIndexedProperty(CarPhone);
            this.RegisterIndexedProperty(CompanyMainPhone);
            this.RegisterIndexedProperty(HomeFax);
            this.RegisterIndexedProperty(HomePhone);
            this.RegisterIndexedProperty(HomePhone2);
            this.RegisterIndexedProperty(Isdn);
            this.RegisterIndexedProperty(MobilePhone);
            this.RegisterIndexedProperty(OtherFax);
            this.RegisterIndexedProperty(OtherTelephone);
            this.RegisterIndexedProperty(Pager);
            this.RegisterIndexedProperty(PrimaryPhone);
            this.RegisterIndexedProperty(RadioPhone);
            this.RegisterIndexedProperty(Telex);
            this.RegisterIndexedProperty(TtyTddPhone);
            this.RegisterIndexedProperty(BusinessAddressStreet);
            this.RegisterIndexedProperty(BusinessAddressCity);
            this.RegisterIndexedProperty(BusinessAddressState);
            this.RegisterIndexedProperty(BusinessAddressCountryOrRegion);
            this.RegisterIndexedProperty(BusinessAddressPostalCode);
            this.RegisterIndexedProperty(HomeAddressStreet);
            this.RegisterIndexedProperty(HomeAddressCity);
            this.RegisterIndexedProperty(HomeAddressState);
            this.RegisterIndexedProperty(HomeAddressCountryOrRegion);
            this.RegisterIndexedProperty(HomeAddressPostalCode);
            this.RegisterIndexedProperty(OtherAddressStreet);
            this.RegisterIndexedProperty(OtherAddressCity);
            this.RegisterIndexedProperty(OtherAddressState);
            this.RegisterIndexedProperty(OtherAddressCountryOrRegion);
            this.RegisterIndexedProperty(OtherAddressPostalCode);
        }

        internal ContactSchema()
            : base()
        {
        }
    }
}
