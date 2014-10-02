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
