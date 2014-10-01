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
