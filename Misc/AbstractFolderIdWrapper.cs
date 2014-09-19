// ---------------------------------------------------------------------------
// <copyright file="AbstractFolderIdWrapper.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the AbstractFolderIdWrapper enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents the abstraction of a folder Id.
    /// </summary>
    internal abstract class AbstractFolderIdWrapper : IJsonSerializable
    {
        /// <summary>
        /// Obtains the Folder object associated with the wrapper.
        /// </summary>
        /// <returns>The Folder object associated with the wrapper.</returns>
        public virtual Folder GetFolder()
        {
            return null;
        }

        /// <summary>
        /// Initializes a new instance of AbstractFolderIdWrapper.
        /// </summary>
        internal AbstractFolderIdWrapper()
        {
        }

        /// <summary>
        /// Writes the Id encapsulated in the wrapper to XML.
        /// </summary>
        /// <param name="writer">The writer to write the Id to.</param>
        internal abstract void WriteToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Validates folderId against specified version.
        /// </summary>
        /// <param name="version">The version.</param>
        internal virtual void Validate(ExchangeVersion version)
        {
        }

        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        object IJsonSerializable.ToJson(ExchangeService service)
        {
            return this.InternalToJson(service);
        }
        
        /// <summary>
        /// Creates a JSON representation of this object.
        /// </summary>
        /// <param name="service">The service.</param>
        /// <returns>
        /// A Json value (either a JsonObject, an array of Json values, or a Json primitive)
        /// </returns>
        internal abstract object InternalToJson(ExchangeService service);
    }
}
