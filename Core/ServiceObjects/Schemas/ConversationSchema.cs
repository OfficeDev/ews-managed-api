// ---------------------------------------------------------------------------
// <copyright file="ConversationSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ConversationSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Diagnostics.CodeAnalysis;

    /// <summary>
    /// Represents the schema for Conversation.
    /// </summary>
    [Schema]
    public class ConversationSchema : ServiceObjectSchema
    {
        /// <summary>
        /// Field URIs for Item.
        /// </summary>
        private static class FieldUris
        {
            public const string ConversationId = "conversation:ConversationId";
            public const string ConversationTopic = "conversation:ConversationTopic";
            public const string UniqueRecipients = "conversation:UniqueRecipients";
            public const string GlobalUniqueRecipients = "conversation:GlobalUniqueRecipients";
            public const string UniqueUnreadSenders = "conversation:UniqueUnreadSenders";
            public const string GlobalUniqueUnreadSenders = "conversation:GlobalUniqueUnreadSenders";
            public const string UniqueSenders = "conversation:UniqueSenders";
            public const string GlobalUniqueSenders = "conversation:GlobalUniqueSenders";
            public const string LastDeliveryTime = "conversation:LastDeliveryTime";
            public const string GlobalLastDeliveryTime = "conversation:GlobalLastDeliveryTime";
            public const string Categories = "conversation:Categories";
            public const string GlobalCategories = "conversation:GlobalCategories";
            public const string FlagStatus = "conversation:FlagStatus";
            public const string GlobalFlagStatus = "conversation:GlobalFlagStatus";
            public const string HasAttachments = "conversation:HasAttachments";
            public const string GlobalHasAttachments = "conversation:GlobalHasAttachments";
            public const string MessageCount = "conversation:MessageCount";
            public const string GlobalMessageCount = "conversation:GlobalMessageCount";
            public const string UnreadCount = "conversation:UnreadCount";
            public const string GlobalUnreadCount = "conversation:GlobalUnreadCount";
            public const string Size = "conversation:Size";
            public const string GlobalSize = "conversation:GlobalSize";
            public const string ItemClasses = "conversation:ItemClasses";
            public const string GlobalItemClasses = "conversation:GlobalItemClasses";
            public const string Importance = "conversation:Importance";
            public const string GlobalImportance = "conversation:GlobalImportance";
            public const string ItemIds = "conversation:ItemIds";
            public const string GlobalItemIds = "conversation:GlobalItemIds";
            public const string LastModifiedTime = "conversation:LastModifiedTime";
            public const string InstanceKey = "conversation:InstanceKey";
            public const string Preview = "conversation:Preview";
            public const string IconIndex = "conversation:IconIndex";
            public const string GlobalIconIndex = "conversation:GlobalIconIndex";
            public const string DraftItemIds = "conversation:DraftItemIds";
            public const string HasIrm = "conversation:HasIrm";
            public const string GlobalHasIrm = "conversation:GlobalHasIrm";
        }

        /// <summary>
        /// Defines the Id property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Id =
            new ComplexPropertyDefinition<ConversationId>(
                XmlElementNames.ConversationId,
                FieldUris.ConversationId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new ConversationId(); });

        /// <summary>
        /// Defines the Topic property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Topic =
            new StringPropertyDefinition(
                XmlElementNames.ConversationTopic,
                FieldUris.ConversationTopic,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the UniqueRecipients property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition UniqueRecipients =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.UniqueRecipients,
                FieldUris.UniqueRecipients,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the GlobalUniqueRecipients property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalUniqueRecipients =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.GlobalUniqueRecipients,
                FieldUris.GlobalUniqueRecipients,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the UniqueUnreadSenders property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition UniqueUnreadSenders =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.UniqueUnreadSenders,
                FieldUris.UniqueUnreadSenders,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the GlobalUniqueUnreadSenders property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalUniqueUnreadSenders =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.GlobalUniqueUnreadSenders,
                FieldUris.GlobalUniqueUnreadSenders,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the UniqueSenders property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition UniqueSenders =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.UniqueSenders,
                FieldUris.UniqueSenders,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the GlobalUniqueSenders property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalUniqueSenders =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.GlobalUniqueSenders,
                FieldUris.GlobalUniqueSenders,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the LastDeliveryTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition LastDeliveryTime =
            new DateTimePropertyDefinition(
                XmlElementNames.LastDeliveryTime,
                FieldUris.LastDeliveryTime,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalLastDeliveryTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalLastDeliveryTime =
            new DateTimePropertyDefinition(
                XmlElementNames.GlobalLastDeliveryTime,
                FieldUris.GlobalLastDeliveryTime,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the Categories property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Categories =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.Categories,
                FieldUris.Categories,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the GlobalCategories property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalCategories =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.GlobalCategories,
                FieldUris.GlobalCategories,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(); });

        /// <summary>
        /// Defines the FlagStatus property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition FlagStatus =
            new GenericPropertyDefinition<ConversationFlagStatus>(
                XmlElementNames.FlagStatus,
                FieldUris.FlagStatus,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalFlagStatus property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalFlagStatus =
            new GenericPropertyDefinition<ConversationFlagStatus>(
                XmlElementNames.GlobalFlagStatus,
                FieldUris.GlobalFlagStatus,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the HasAttachments property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition HasAttachments =
            new BoolPropertyDefinition(
                XmlElementNames.HasAttachments,
                FieldUris.HasAttachments,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalHasAttachments property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalHasAttachments =
            new BoolPropertyDefinition(
                XmlElementNames.GlobalHasAttachments,
                FieldUris.GlobalHasAttachments,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the MessageCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition MessageCount =
            new IntPropertyDefinition(
                XmlElementNames.MessageCount,
                FieldUris.MessageCount,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalMessageCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalMessageCount =
            new IntPropertyDefinition(
                XmlElementNames.GlobalMessageCount,
                FieldUris.GlobalMessageCount,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the UnreadCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition UnreadCount =
            new IntPropertyDefinition(
                XmlElementNames.UnreadCount,
                FieldUris.UnreadCount,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalUnreadCount property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalUnreadCount =
            new IntPropertyDefinition(
                XmlElementNames.GlobalUnreadCount,
                FieldUris.GlobalUnreadCount,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

                /// <summary>
        /// Defines the Size property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Size =
            new IntPropertyDefinition(
                XmlElementNames.Size,
                FieldUris.Size,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalSize property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalSize =
            new IntPropertyDefinition(
                XmlElementNames.GlobalSize,
                FieldUris.GlobalSize,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the ItemClasses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ItemClasses =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.ItemClasses,
                FieldUris.ItemClasses,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(XmlElementNames.ItemClass); });

        /// <summary>
        /// Defines the GlobalItemClasses property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalItemClasses =
            new ComplexPropertyDefinition<StringList>(
                XmlElementNames.GlobalItemClasses,
                FieldUris.GlobalItemClasses,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new StringList(XmlElementNames.ItemClass); });

        /// <summary>
        /// Defines the Importance property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Importance =
            new GenericPropertyDefinition<Importance>(
                XmlElementNames.Importance,
                FieldUris.Importance,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the GlobalImportance property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalImportance =
            new GenericPropertyDefinition<Importance>(
                XmlElementNames.GlobalImportance,
                FieldUris.GlobalImportance,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1);

        /// <summary>
        /// Defines the ItemIds property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ItemIds =
            new ComplexPropertyDefinition<ItemIdCollection>(
                XmlElementNames.ItemIds,
                FieldUris.ItemIds,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new ItemIdCollection(); });

        /// <summary>
        /// Defines the GlobalItemIds property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalItemIds =
            new ComplexPropertyDefinition<ItemIdCollection>(
                XmlElementNames.GlobalItemIds,
                FieldUris.GlobalItemIds,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2010_SP1,
                delegate() { return new ItemIdCollection(); });

        /// <summary>
        /// Defines the LastModifiedTime property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition LastModifiedTime =
            new DateTimePropertyDefinition(
                XmlElementNames.LastModifiedTime,
                FieldUris.LastModifiedTime,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// Defines the InstanceKey property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition InstanceKey =
            new ByteArrayPropertyDefinition(
                XmlElementNames.InstanceKey,
                FieldUris.InstanceKey,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// Defines the Preview property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Preview =
            new StringPropertyDefinition(
                XmlElementNames.Preview,
                FieldUris.Preview,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// Defines the IconIndex property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IconIndex =
            new GenericPropertyDefinition<IconIndex>(
                XmlElementNames.IconIndex,
                FieldUris.IconIndex,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// Defines the GlobalIconIndex property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalIconIndex =
            new GenericPropertyDefinition<IconIndex>(
                XmlElementNames.GlobalIconIndex,
                FieldUris.GlobalIconIndex,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// Defines the DraftItemIds property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition DraftItemIds =
            new ComplexPropertyDefinition<ItemIdCollection>(
                XmlElementNames.DraftItemIds,
                FieldUris.DraftItemIds,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013,
                delegate { return new ItemIdCollection(); });

        /// <summary>
        /// Defines the HasIrm property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition HasIrm =
            new BoolPropertyDefinition(
                XmlElementNames.HasIrm,
                FieldUris.HasIrm,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        /// <summary>
        /// Defines the GlobalHasIrm property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition GlobalHasIrm =
            new BoolPropertyDefinition(
                XmlElementNames.GlobalHasIrm,
                FieldUris.GlobalHasIrm,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2013);

        // This must be declared after the property definitions
        internal static readonly ConversationSchema Instance = new ConversationSchema();

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
            this.RegisterProperty(Topic);
            this.RegisterProperty(UniqueRecipients);
            this.RegisterProperty(GlobalUniqueRecipients);
            this.RegisterProperty(UniqueUnreadSenders);
            this.RegisterProperty(GlobalUniqueUnreadSenders);
            this.RegisterProperty(UniqueSenders);
            this.RegisterProperty(GlobalUniqueSenders);
            this.RegisterProperty(LastDeliveryTime);
            this.RegisterProperty(GlobalLastDeliveryTime);
            this.RegisterProperty(Categories);
            this.RegisterProperty(GlobalCategories);
            this.RegisterProperty(FlagStatus);
            this.RegisterProperty(GlobalFlagStatus);
            this.RegisterProperty(HasAttachments);
            this.RegisterProperty(GlobalHasAttachments);
            this.RegisterProperty(MessageCount);
            this.RegisterProperty(GlobalMessageCount);
            this.RegisterProperty(UnreadCount);
            this.RegisterProperty(GlobalUnreadCount);
            this.RegisterProperty(Size);
            this.RegisterProperty(GlobalSize);
            this.RegisterProperty(ItemClasses);
            this.RegisterProperty(GlobalItemClasses);
            this.RegisterProperty(Importance);
            this.RegisterProperty(GlobalImportance);
            this.RegisterProperty(ItemIds);
            this.RegisterProperty(GlobalItemIds);
            this.RegisterProperty(LastModifiedTime);
            this.RegisterProperty(InstanceKey);
            this.RegisterProperty(Preview);
            this.RegisterProperty(IconIndex);
            this.RegisterProperty(GlobalIconIndex);
            this.RegisterProperty(DraftItemIds);
            this.RegisterProperty(HasIrm);
            this.RegisterProperty(GlobalHasIrm);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationSchema"/> class.
        /// </summary>
        internal ConversationSchema()
            : base()
        {
        }
    }
}