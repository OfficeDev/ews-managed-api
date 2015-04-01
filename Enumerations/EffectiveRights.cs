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
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the effective user rights associated with an item or folder.
    /// </summary>
    [Flags]
    public enum EffectiveRights
    {
        /// <summary>
        /// The user has no acces right on the item or folder.
        /// </summary>
        None = 0,

        /// <summary>
        /// The user can create associated items (FAI)
        /// </summary>
        CreateAssociated = 1,

        /// <summary>
        /// The user can create items.
        /// </summary>
        CreateContents = 2,

        /// <summary>
        /// The user can create sub-folders.
        /// </summary>
        CreateHierarchy = 4,

        /// <summary>
        /// The user can delete items and/or folders.
        /// </summary>
        Delete = 8,

        /// <summary>
        /// The user can modify the properties of items and/or folders.
        /// </summary>
        Modify = 16,

        /// <summary>
        /// The user can read the contents of items.
        /// </summary>
        Read = 32,

        /// <summary>
        /// The user can view private items.
        /// </summary>
        ViewPrivateItems = 64
    }
}