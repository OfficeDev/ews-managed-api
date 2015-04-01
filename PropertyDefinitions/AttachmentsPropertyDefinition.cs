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