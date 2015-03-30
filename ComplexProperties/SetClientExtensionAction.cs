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
    using System;

    /// <summary>
    /// Represents the SetClientExtension method action.
    /// </summary>
    public sealed class SetClientExtensionAction : ComplexProperty
    {
        private readonly SetClientExtensionActionId setClientExtensionActionId;
        private readonly string extensionId;
        private readonly ClientExtension clientExtension;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetClientExtensionAction"/> class.
        /// </summary>
        /// <param name="setClientExtensionActionId">Set action such as install, uninstall and configure</param>
        /// <param name="extensionId">ExtensionId, required by configure and uninstall actions</param>
        /// <param name="clientExtension">Extension data object, e.g. required by configure action</param>
        public SetClientExtensionAction(
            SetClientExtensionActionId setClientExtensionActionId,
            string extensionId,
            ClientExtension clientExtension)
                : base()
        {
            this.Namespace = XmlNamespace.Types;
            this.setClientExtensionActionId = setClientExtensionActionId;
            this.extensionId = extensionId;
            this.clientExtension = clientExtension;
        }

        /// <summary>
        /// Writes attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
        {
            writer.WriteAttributeValue(XmlAttributeNames.SetClientExtensionActionId, this.setClientExtensionActionId);

            if (!string.IsNullOrEmpty(this.extensionId))
            {
                writer.WriteAttributeValue(XmlAttributeNames.ClientExtensionId, this.extensionId);
            }
        }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (null != this.clientExtension)
            {
                this.clientExtension.WriteToXml(writer, XmlNamespace.Types, XmlElementNames.ClientExtension);
            }
        }
    }
}