// ---------------------------------------------------------------------------
// <copyright file="PostReplySchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PostReplySchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents PostReply schema definition.
    /// </summary>
    internal sealed class PostReplySchema : ServiceObjectSchema
    {
        // This must be declared after the property definitions
        internal static readonly PostReplySchema Instance = new PostReplySchema();

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
            this.RegisterProperty(ResponseObjectSchema.ReferenceItemId);
            this.RegisterProperty(ResponseObjectSchema.BodyPrefix);
        }
    }
}
