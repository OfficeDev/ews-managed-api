// ---------------------------------------------------------------------------
// <copyright file="FileAsMapping.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FileAsMapping enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Defines the way the FileAs property of a contact is automatically formatted.
    /// </summary>
    public enum FileAsMapping
    {
        /// <summary>
        /// No automatic formatting is used.
        /// </summary>
        None,

        /// <summary>
        /// Surname, GivenName
        /// </summary>
        [EwsEnum("LastCommaFirst")]
        SurnameCommaGivenName,

        /// <summary>
        /// GivenName Surname
        /// </summary>
        [EwsEnum("FirstSpaceLast")]
        GivenNameSpaceSurname,

        /// <summary>
        /// Company
        /// </summary>
        Company,

        /// <summary>
        /// Surname, GivenName (Company)
        /// </summary>
        [EwsEnum("LastCommaFirstCompany")]
        SurnameCommaGivenNameCompany,

        /// <summary>
        /// Company (SurnameGivenName)
        /// </summary>
        [EwsEnum("CompanyLastFirst")]
        CompanySurnameGivenName,

        /// <summary>
        /// SurnameGivenName
        /// </summary>
        [EwsEnum("LastFirst")]
        SurnameGivenName,

        /// <summary>
        /// SurnameGivenName (Company)
        /// </summary>
        [EwsEnum("LastFirstCompany")]
        SurnameGivenNameCompany,

        /// <summary>
        /// Company (Surname, GivenName)
        /// </summary>
        [EwsEnum("CompanyLastCommaFirst")]
        CompanySurnameCommaGivenName,

        /// <summary>
        /// SurnameGivenName Suffix
        /// </summary>
        [EwsEnum("LastFirstSuffix")]
        SurnameGivenNameSuffix,

        /// <summary>
        /// Surname GivenName (Company)
        /// </summary>
        [EwsEnum("LastSpaceFirstCompany")]
        SurnameSpaceGivenNameCompany,

        /// <summary>
        /// Company (Surname GivenName)
        /// </summary>
        [EwsEnum("CompanyLastSpaceFirst")]
        CompanySurnameSpaceGivenName,

        /// <summary>
        /// Surname GivenName
        /// </summary>
        [EwsEnum("LastSpaceFirst")]
        SurnameSpaceGivenName,

        /// <summary>
        /// Display Name (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        DisplayName,

        /// <summary>
        /// GivenName (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        [EwsEnum("FirstName")]
        GivenName,

        /// <summary>
        /// Surname GivenName Middle Suffix (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        [EwsEnum("LastFirstMiddleSuffix")]
        SurnameGivenNameMiddleSuffix,

        /// <summary>
        /// Surname (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        [EwsEnum("LastName")]
        Surname,

        /// <summary>
        /// Empty (Exchange 2010 or later).
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2010)]
        Empty
    }
}
