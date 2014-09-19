// ---------------------------------------------------------------------------
// <copyright file="ResponseMessageSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseMessageSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents ResponseMessage schema definition.
    /// </summary>
    internal class ResponseMessageSchema : ServiceObjectSchema
    {
        // This must be declared after the property definitions
        internal static readonly ResponseMessageSchema Instance = new ResponseMessageSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(ItemSchema.Subject);
            this.RegisterProperty(ItemSchema.Body);
            this.RegisterProperty(EmailMessageSchema.ToRecipients);
            this.RegisterProperty(EmailMessageSchema.CcRecipients);
            this.RegisterProperty(EmailMessageSchema.BccRecipients);
            this.RegisterProperty(EmailMessageSchema.IsReadReceiptRequested);
            this.RegisterProperty(EmailMessageSchema.IsDeliveryReceiptRequested);
            this.RegisterProperty(ResponseObjectSchema.ReferenceItemId);
            this.RegisterProperty(ResponseObjectSchema.BodyPrefix);
        }
    }
}
