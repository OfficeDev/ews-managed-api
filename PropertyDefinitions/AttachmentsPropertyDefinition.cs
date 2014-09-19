// ---------------------------------------------------------------------------
// <copyright file="AttachmentsPropertyDefinition.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents base Attachments property type.
    /// </summary>
    internal sealed class AttachmentsPropertyDefinition : ComplexPropertyDefinition<AttachmentCollection>
    {
        private static readonly PropertyDefinitionFlags Exchange2010SP2PropertyDefinitionFlags =
            PropertyDefinitionFlags.AutoInstantiateOnRead |
            PropertyDefinitionFlags.CanSet |
            PropertyDefinitionFlags.ReuseInstance |
            PropertyDefinitionFlags.UpdateCollectionItems;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachmentsPropertyDefinition"/> class.
        /// </summary>
        public AttachmentsPropertyDefinition() :
            base(
            XmlElementNames.Attachments,
            "item:Attachments",
            PropertyDefinitionFlags.AutoInstantiateOnRead,
            ExchangeVersion.Exchange2007_SP1,
            delegate() { return new AttachmentCollection(); })
        {
        }

        /// <summary>
        /// Determines whether the specified flag is set.
        /// </summary>
        /// <param name="flag">The flag.</param>
        /// <param name="version">Requested version.</param>
        /// <returns>
        ///     <c>true</c> if the specified flag is set; otherwise, <c>false</c>.
        /// </returns>
        internal override bool HasFlag(PropertyDefinitionFlags flag, ExchangeVersion? version)
        {
            if (version != null && version >= ExchangeVersion.Exchange2010_SP2)
            {
                return (flag & AttachmentsPropertyDefinition.Exchange2010SP2PropertyDefinitionFlags) == flag;
            }

            return base.HasFlag(flag, version);
        }
    }
}