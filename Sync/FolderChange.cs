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
// <summary>Defines the FolderChange class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents a change on a folder as returned by a synchronization operation.
    /// </summary>
    public sealed class FolderChange : Change
    {
        /// <summary>
        /// Initializes a new instance of FolderChange.
        /// </summary>
        internal FolderChange()
            : base()
        {
        }

        /// <summary>
        /// Creates a FolderId instance.
        /// </summary>
        /// <returns>A FolderId.</returns>
        internal override ServiceId CreateId()
        {
            return new FolderId();
        }

        /// <summary>
        /// Gets the folder the change applies to. Folder is null when ChangeType is equal to
        /// ChangeType.Delete. In that case, use the FolderId property to retrieve the Id of
        /// the folder that was deleted.
        /// </summary>
        public Folder Folder
        {
            get { return (Folder)this.ServiceObject; }
        }

        /// <summary>
        /// Gets the Id of the folder the change applies to.
        /// </summary>
        public FolderId FolderId
        {
            get { return (FolderId)this.Id; }
        }
    }
}
