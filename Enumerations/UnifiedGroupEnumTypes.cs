// ---------------------------------------------------------------------------
// <copyright file="UnifiedGroupsEnumTypes.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the UnifiedGroupsEnumTypes.</summary>
//-----------------------------------------------------------------------
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
