// ---------------------------------------------------------------------------
// <copyright file="PostItemSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PostItemSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for post items.
    /// </summary>
    [Schema]
    public sealed class PostItemSchema : ItemSchema
    {
        /// <summary>
        /// Field URIs for PostItem.
        /// </summary>
        private static class FieldUris
        {
            public const string PostedTime = "postitem:PostedTime";
        }

        /// <summary>
        /// Defines the ConversationIndex property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConversationIndex =
            EmailMessageSchema.ConversationIndex;

        /// <summary>
        /// Defines the ConversationTopic property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConversationTopic =
            EmailMessageSchema.ConversationTopic;

        /// <summary>
        /// Defines the From property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition From =
            EmailMessageSchema.From;

        /// <summary>
        /// Defines the InternetMessageId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition InternetMessageId =
            EmailMessageSchema.InternetMessageId;

        /// <summary>
        /// Defines the IsRead property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsRead =
            EmailMessageSchema.IsRead;

        /// <summary>
        /// Defines the PostedTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition PostedTime =
            new DateTimePropertyDefinition(
                XmlElementNames.PostedTime,
                FieldUris.PostedTime,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the References property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition References =
            EmailMessageSchema.References;

        /// <summary>
        /// Defines the Sender property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Sender =
            EmailMessageSchema.Sender;

        // This must be after the declaration of property definitions
        internal static new readonly PostItemSchema Instance = new PostItemSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(ConversationIndex);
            this.RegisterProperty(ConversationTopic);
            this.RegisterProperty(From);
            this.RegisterProperty(InternetMessageId);
            this.RegisterProperty(IsRead);
            this.RegisterProperty(PostedTime);
            this.RegisterProperty(References);
            this.RegisterProperty(Sender);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PostItemSchema"/> class.
        /// </summary>
        internal PostItemSchema()
            : base()
        {
        }
    }
}
