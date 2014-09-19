// ---------------------------------------------------------------------------
// <copyright file="CancelMeetingMessageSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CancelMeetingMessageSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents CancelMeetingMessage schema definition.
    /// </summary>
    internal class CancelMeetingMessageSchema : ServiceObjectSchema
    {
        public static readonly PropertyDefinition Body =
            new ComplexPropertyDefinition<MessageBody>(
                XmlElementNames.NewBodyContent,
                PropertyDefinitionFlags.CanSet,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new MessageBody(); });

        // This must be declared after the property definitions
        internal static readonly CancelMeetingMessageSchema Instance = new CancelMeetingMessageSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(EmailMessageSchema.IsReadReceiptRequested);
            this.RegisterProperty(EmailMessageSchema.IsDeliveryReceiptRequested);
            this.RegisterProperty(ResponseObjectSchema.ReferenceItemId);
            this.RegisterProperty(CancelMeetingMessageSchema.Body);
        }
    }
}
