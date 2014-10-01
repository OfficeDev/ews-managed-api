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
// <summary>Defines the FolderSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for folders.
    /// </summary>
    [Schema]
    public class FolderSchema : ServiceObjectSchema
    {
        /// <summary>
        /// Field URIs for folders.
        /// </summary>
        private static class FieldUris
        {
            public const string FolderId = "folder:FolderId";
            public const string ParentFolderId = "folder:ParentFolderId";
            public const string DisplayName = "folder:DisplayName";
            public const string UnreadCount = "folder:UnreadCount";
            public const string TotalCount = "folder:TotalCount";
            public const string ChildFolderCount = "folder:ChildFolderCount";
            public const string FolderClass = "folder:FolderClass";
            public const string ManagedFolderInformation = "folder:ManagedFolderInformation";
            public const string EffectiveRights = "folder:EffectiveRights";
            public const string PermissionSet = "folder:PermissionSet";
            public const string PolicyTag = "folder:PolicyTag";
            public const string ArchiveTag = "folder:ArchiveTag";
            public const string DistinguishedFolderId = "folder:DistinguishedFolderId";
        }

        /// <summary>
        /// Defines the Id property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Id =
            new ComplexPropertyDefinition<FolderId>(
                XmlElementNames.FolderId,
                FieldUris.FolderId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new FolderId(); });

        /// <summary>
        /// Defines the FolderClass property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FolderClass =
            new StringPropertyDefinition(
                XmlElementNames.FolderClass,
                FieldUris.FolderClass,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ParentFolderId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ParentFolderId =
            new ComplexPropertyDefinition<FolderId>(
                XmlElementNames.ParentFolderId,
                FieldUris.ParentFolderId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new FolderId(); });

        /// <summary>
        /// Defines the ChildFolderCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ChildFolderCount =
            new IntPropertyDefinition(
                XmlElementNames.ChildFolderCount,
                FieldUris.ChildFolderCount,
                PropertyDefinitionFlags.CanFind,
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
        /// Defines the UnreadCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition UnreadCount =
            new IntPropertyDefinition(
                XmlElementNames.UnreadCount,
                FieldUris.UnreadCount,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the TotalCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition TotalCount =
            new IntPropertyDefinition(
                XmlElementNames.TotalCount,
                FieldUris.TotalCount,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ManagedFolderInformation property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ManagedFolderInformation =
            new ComplexPropertyDefinition<ManagedFolderInformation>(
                XmlElementNames.ManagedFolderInformation,
                FieldUris.ManagedFolderInformation,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new ManagedFolderInformation(); });

        /// <summary>
        /// Defines the EffectiveRights property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition EffectiveRights =
            new EffectiveRightsPropertyDefinition(
                XmlElementNames.EffectiveRights,
                FieldUris.EffectiveRights,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the Permissions property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Permissions =
            new PermissionSetPropertyDefinition(
                XmlElementNames.PermissionSet,
                FieldUris.PermissionSet,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.MustBeExplicitlyLoaded,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the WellKnownFolderName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition WellKnownFolderName =
            new GenericPropertyDefinition<WellKnownFolderName>(
                XmlElementNames.DistinguishedFolderId,
                FieldUris.DistinguishedFolderId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                true);

        /// <summary>
        /// Defines the PolicyTag property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PolicyTag =
            new ComplexPropertyDefinition<PolicyTag>(
                XmlElementNames.PolicyTag,
                FieldUris.PolicyTag,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate() { return new PolicyTag(); });

        /// <summary>
        /// Defines the ArchiveTag property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ArchiveTag =
            new ComplexPropertyDefinition<ArchiveTag>(
                XmlElementNames.ArchiveTag,
                FieldUris.ArchiveTag,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate() { return new ArchiveTag(); });

        // This must be declared after the property definitions
        internal static readonly FolderSchema Instance = new FolderSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(Id);
            this.RegisterProperty(ParentFolderId);
            this.RegisterProperty(FolderClass);
            this.RegisterProperty(DisplayName);
            this.RegisterProperty(TotalCount);
            this.RegisterProperty(ChildFolderCount);
            this.RegisterProperty(ServiceObjectSchema.ExtendedProperties);
            this.RegisterProperty(ManagedFolderInformation);
            this.RegisterProperty(EffectiveRights);
            this.RegisterProperty(Permissions);
            this.RegisterProperty(UnreadCount);
            this.RegisterProperty(WellKnownFolderName);
            this.RegisterProperty(PolicyTag);
            this.RegisterProperty(ArchiveTag);
        }
    }
}