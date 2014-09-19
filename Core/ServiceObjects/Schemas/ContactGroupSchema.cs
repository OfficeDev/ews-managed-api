// ---------------------------------------------------------------------------
// <copyright file="ContactGroupSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ContactGroupSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for contact groups.
    /// </summary>
    [Schema]
    public class ContactGroupSchema : ItemSchema
    {
        /// <summary>
        /// Defines the DisplayName property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DisplayName =
            ContactSchema.DisplayName;

        /// <summary>
        /// Defines the FileAs property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FileAs =
            ContactSchema.FileAs;

        /// <summary>
        /// Defines the Members property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Members =
            new ComplexPropertyDefinition<GroupMemberCollection>(
                XmlElementNames.Members,
                FieldUris.Members,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate,
                ExchangeVersion.Exchange2010,
                delegate() { return new GroupMemberCollection(); });

        /// <summary>
        /// This must be declared after the property definitions.
        /// </summary>
        internal static new readonly ContactGroupSchema Instance = new ContactGroupSchema();

        /// <summary>
        /// Initializes a new instance of the <see cref="ContactGroupSchema"/> class.
        /// </summary>
        internal ContactGroupSchema()
            : base()
        {
        }

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(DisplayName);
            this.RegisterProperty(FileAs);
            this.RegisterProperty(Members);
        }

        /// <summary>
        /// Field URIs for Members.
        /// </summary>
        private static class FieldUris
        {
            /// <summary>
            /// FieldUri for members.
            /// </summary>
            public const string Members = "distributionlist:Members";
        }
    }
}