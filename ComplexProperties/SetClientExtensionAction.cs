// ---------------------------------------------------------------------------
// <copyright file="SetClientExtensionAction.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the SetClientExtensionAction class.</summary>
//-----------------------------------------------------------------------

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
