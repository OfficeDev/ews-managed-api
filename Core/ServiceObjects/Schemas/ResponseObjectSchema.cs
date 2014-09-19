// ---------------------------------------------------------------------------
// <copyright file="ResponseObjectSchema.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ResponseObjectSchema class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents ResponseObject schema definition.
    /// </summary>
    internal class ResponseObjectSchema : ServiceObjectSchema
    {
        public static readonly PropertyDefinition ReferenceItemId =
            new ComplexPropertyDefinition<ItemId>(
                XmlElementNames.ReferenceItemId,
                PropertyDefinitionFlags.AutoInstantiateOnRead | PropertyDefinitionFlags.CanSet,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new ItemId(); });

        public static readonly PropertyDefinition BodyPrefix =
            new ComplexPropertyDefinition<MessageBody>(
                XmlElementNames.NewBodyContent,
                PropertyDefinitionFlags.CanSet,
                ExchangeVersion.Exchange2007_SP1,
                delegate() { return new MessageBody(); });

        // This must be declared after the property definitions
        internal static readonly ResponseObjectSchema Instance = new ResponseObjectSchema();

        /// <summary>
        /// Registers properties.
        /// </summary>
        /// <remarks>
        /// IMPORTANT NOTE: PROPERTIES MUST BE REGISTERED IN SCHEMA ORDER (i.e. the same order as they are defined in types.xsd)
        /// </remarks>
        internal override void RegisterProperties()
        {
            base.RegisterProperties();

            this.RegisterProperty(ResponseObjectSchema.ReferenceItemId);
        }
    }
}
