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
    /// Represents the schema for e-mail messages.
    /// </summary>
    [Schema]
    public class EmailMessageSchema : ItemSchema
    {
        /// <summary>
        /// Field URIs for EmailMessage.
        /// </summary>
        private static class FieldUris
        {
            public const string ConversationIndex = "message:ConversationIndex";
            public const string ConversationTopic = "message:ConversationTopic";
            public const string InternetMessageId = "message:InternetMessageId";
            public const string IsRead = "message:IsRead";
            public const string IsResponseRequested = "message:IsResponseRequested";
            public const string IsReadReceiptRequested = "message:IsReadReceiptRequested";
            public const string IsDeliveryReceiptRequested = "message:IsDeliveryReceiptRequested";
            public const string References = "message:References";
            public const string ReplyTo = "message:ReplyTo";
            public const string From = "message:From";
            public const string Sender = "message:Sender";
            public const string ToRecipients = "message:ToRecipients";
            public const string CcRecipients = "message:CcRecipients";
            public const string BccRecipients = "message:BccRecipients";
            public const string ReceivedBy = "message:ReceivedBy";
            public const string ReceivedRepresenting = "message:ReceivedRepresenting";
            public const string ApprovalRequestData = "message:ApprovalRequestData";
            public const string VotingInformation = "message:VotingInformation";
            public const string Likers = "message:Likers";
        }

        /// <summary>
        /// Defines the ToRecipients property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ToRecipients =
            new ComplexPropertyDefinition<EmailAddressCollection>(
                XmlElementNames.ToRecipients,
                FieldUris.ToRecipients,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddressCollection(); });

        /// <summary>
        /// Defines the BccRecipients property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition BccRecipients =
            new ComplexPropertyDefinition<EmailAddressCollection>(
                XmlElementNames.BccRecipients,
                FieldUris.BccRecipients,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddressCollection(); });

        /// <summary>
        /// Defines the CcRecipients property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition CcRecipients =
            new ComplexPropertyDefinition<EmailAddressCollection>(
                XmlElementNames.CcRecipients,
                FieldUris.CcRecipients,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddressCollection(); });

        /// <summary>
        /// Defines the ConversationIndex property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConversationIndex =
            new ByteArrayPropertyDefinition(
                XmlElementNames.ConversationIndex,
                FieldUris.ConversationIndex,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ConversationTopic property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ConversationTopic =
            new StringPropertyDefinition(
                XmlElementNames.ConversationTopic,
                FieldUris.ConversationTopic,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the From property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition From =
            new ContainedPropertyDefinition<EmailAddress>(
                XmlElementNames.From,
                FieldUris.From,
                XmlElementNames.Mailbox,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddress(); });

        /// <summary>
        /// Defines the IsDeliveryReceiptRequested property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsDeliveryReceiptRequested =
            new BoolPropertyDefinition(
                XmlElementNames.IsDeliveryReceiptRequested,
                FieldUris.IsDeliveryReceiptRequested,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsRead property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsRead =
            new BoolPropertyDefinition(
                XmlElementNames.IsRead,
                FieldUris.IsRead,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsReadReceiptRequested property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsReadReceiptRequested =
            new BoolPropertyDefinition(
                XmlElementNames.IsReadReceiptRequested,
                FieldUris.IsReadReceiptRequested,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the IsResponseRequested property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition IsResponseRequested =
            new BoolPropertyDefinition(
                XmlElementNames.IsResponseRequested,
                FieldUris.IsResponseRequested,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                true);  // isNullable

        /// <summary>
        /// Defines the InternetMessageId property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition InternetMessageId =
            new StringPropertyDefinition(
                XmlElementNames.InternetMessageId,
                FieldUris.InternetMessageId,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the References property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition References =
            new StringPropertyDefinition(
                XmlElementNames.References,
                FieldUris.References,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1);

        /// <summary>
        /// Defines the ReplyTo property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ReplyTo =
            new ComplexPropertyDefinition<EmailAddressCollection>(
                XmlElementNames.ReplyTo,
                FieldUris.ReplyTo,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanUpdate | PropertyDefinitionFlags.CanDelete,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddressCollection(); });

        /// <summary>
        /// Defines the Sender property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Sender =
            new ContainedPropertyDefinition<EmailAddress>(
                XmlElementNames.Sender,
                FieldUris.Sender,
                XmlElementNames.Mailbox,
                PropertyDefinitionFlags.CanSet | PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddress(); });

        /// <summary>
        /// Defines the ReceivedBy property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ReceivedBy =
            new ContainedPropertyDefinition<EmailAddress>(
                XmlElementNames.ReceivedBy,
                FieldUris.ReceivedBy,
                XmlElementNames.Mailbox,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddress(); });

        /// <summary>
        /// Defines the ReceivedRepresenting property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ReceivedRepresenting =
            new ContainedPropertyDefinition<EmailAddress>(
                XmlElementNames.ReceivedRepresenting,
                FieldUris.ReceivedRepresenting,
                XmlElementNames.Mailbox,
                PropertyDefinitionFlags.CanFind,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new EmailAddress(); });

        /// <summary>
        /// Defines the ApprovalRequestData property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition ApprovalRequestData =
            new ComplexPropertyDefinition<ApprovalRequestData>(
                XmlElementNames.ApprovalRequestData,
                FieldUris.ApprovalRequestData,
                ExchangeVersion.Exchange2013,
                delegate() { return new ApprovalRequestData(); });

        /// <summary>
        /// Defines the VotingInformation property.
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition VotingInformation =
            new ComplexPropertyDefinition<VotingInformation>(
                XmlElementNames.VotingInformation,
                FieldUris.VotingInformation,
                ExchangeVersion.Exchange2013,
                delegate() { return new VotingInformation(); });

        /// <summary>
        /// Defines the Likers property
        /// </summary>
        [SuppressMessage("Microsoft.Security", "CA2104:DoNotDeclareReadOnlyMutableReferenceTypes", Justification = "Immutable type")]
        public static readonly PropertyDefinition Likers =
            new ComplexPropertyDefinition<EmailAddressCollection>(
                XmlElementNames.Likers,
                FieldUris.Likers,
                PropertyDefinitionFlags.AutoInstantiateOnRead,
                ExchangeVersion.Exchange2016,
                delegate() { return new EmailAddressCollection(); });

        // This must be after the declaration of property definitions
        internal static new readonly EmailMessageSchema Instance = new EmailMessageSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(Sender);
            this.RegisterProperty(ToRecipients);
            this.RegisterProperty(CcRecipients);
            this.RegisterProperty(BccRecipients);
            this.RegisterProperty(IsReadReceiptRequested);
            this.RegisterProperty(IsDeliveryReceiptRequested);
            this.RegisterProperty(ConversationIndex);
            this.RegisterProperty(ConversationTopic);
            this.RegisterProperty(From);
            this.RegisterProperty(InternetMessageId);
            this.RegisterProperty(IsRead);
            this.RegisterProperty(IsResponseRequested);
            this.RegisterProperty(References);
            this.RegisterProperty(ReplyTo);
            this.RegisterProperty(ReceivedBy);
            this.RegisterProperty(ReceivedRepresenting);
            this.RegisterProperty(ApprovalRequestData);
            this.RegisterProperty(VotingInformation);
            this.RegisterProperty(Likers);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailMessageSchema"/> class.
        /// </summary>
        internal EmailMessageSchema()
            : base()
        {
        }
    }
}