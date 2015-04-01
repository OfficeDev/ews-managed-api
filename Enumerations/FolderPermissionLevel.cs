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

    // TODO : Do we want to include more information about what those levels actually allow users to do?

    /// <summary>
    /// Defines permission levels for calendar folders.
    /// </summary>
    public enum FolderPermissionLevel
    {
        /// <summary>
        /// No permission is granted.
        /// </summary>
        None,

        /// <summary>
        /// The Owner level.
        /// </summary>
        Owner,

        /// <summary>
        /// The Publishing Editor level.
        /// </summary>
        PublishingEditor,

        /// <summary>
        /// The Editor level.
        /// </summary>
        Editor,

        /// <summary>
        /// The Pusnlishing Author level.
        /// </summary>
        PublishingAuthor,

        /// <summary>
        /// The Author level.
        /// </summary>
        Author,

        /// <summary>
        /// The Non-editing Author level.
        /// </summary>
        NoneditingAuthor,

        /// <summary>
        /// The Reviewer level.
        /// </summary>
        Reviewer,

        /// <summary>
        /// The Contributor level.
        /// </summary>
        Contributor,

        /// <summary>
        /// The Free/busy Time Only level. (Can only be applied to Calendar folders).
        /// </summary>
        FreeBusyTimeOnly,

        /// <summary>
        /// The Free/busy Time, Subject and Location level. (Can only be applied to Calendar folders).
        /// </summary>
        FreeBusyTimeAndSubjectAndLocation,

        /// <summary>
        /// The Custom level.
        /// </summary>
        Custom
    }
}