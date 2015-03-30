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

namespace Microsoft.Exchange.WebServices.Data.Groups
{
    using System;

    /// <summary>
    /// The UnifiedGroupsSortType enum
    /// </summary>
    public enum UnifiedGroupsSortType
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,

        /// <summary>
        /// Display Name
        /// </summary>
        DisplayName = 1,

        /// <summary>
        /// Join Date
        /// </summary>
        JoinDate = 2,

        /// <summary>
        /// Favorite Date
        /// </summary>
        FavoriteDate = 3,

        /// <summary>
        /// Relevance
        /// </summary>
        Relevance = 4,
    }

    /// <summary>
    /// The UnifiedGroupsFilterType enum
    /// </summary>
    public enum UnifiedGroupsFilterType
    {
        /// <summary>
        /// All
        /// </summary>
        All = 0,

        /// <summary>
        /// Favorites
        /// </summary>
        Favorites = 1,

        /// <summary>
        /// Exclude Favorites
        /// </summary>
        ExcludeFavorites = 2
    }

    /// <summary>
    /// The UnifiedGroupAccessType enum
    /// </summary>
    public enum UnifiedGroupAccessType
    {
        /// <summary>
        /// None 
        /// </summary>
        None = 0,

        /// <summary>
        /// Private Group
        /// </summary>
        Private = 1,

        /// <summary>
        /// Secret Group
        /// </summary>
        Secret = 2,

        /// <summary>
        /// Public Group
        /// </summary>
        Public = 3,
    }

    /// <summary>
    /// The UnifiedGroupIdentityType enum
    /// </summary>
    public enum UnifiedGroupIdentityType
    {
        /// <summary>
        /// Smtp Address
        /// </summary>
        SmtpAddress = 0,

        /// <summary>
        /// Legacy DN
        /// </summary>
        LegacyDn = 1,

        /// <summary>
        /// ExternalDirectoryObjectId
        /// </summary>
        ExternalDirectoryObjectId = 2
    }
}