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