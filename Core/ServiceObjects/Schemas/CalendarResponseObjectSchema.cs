// ---------------------------------------------------------------------------
// <copyright file="CalendarResponseObjectSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CalendarResponseObjectSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    internal class CalendarResponseObjectSchema : ServiceObjectSchema
    {
        // This must be declared after the property definitions
        internal static readonly CalendarResponseObjectSchema Instance = new CalendarResponseObjectSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(ItemSchema.ItemClass);
            this.RegisterProperty(ItemSchema.Sensitivity);
            this.RegisterProperty(ItemSchema.Body);
            this.RegisterProperty(ItemSchema.Attachments);
            this.RegisterProperty(ItemSchema.InternetMessageHeaders);
            this.RegisterProperty(EmailMessageSchema.Sender);
            this.RegisterProperty(EmailMessageSchema.ToRecipients);
            this.RegisterProperty(EmailMessageSchema.CcRecipients);
            this.RegisterProperty(EmailMessageSchema.BccRecipients);
            this.RegisterProperty(EmailMessageSchema.IsReadReceiptRequested);
            this.RegisterProperty(EmailMessageSchema.IsDeliveryReceiptRequested);
            this.RegisterProperty(ResponseObjectSchema.ReferenceItemId);
        }
    }
}
