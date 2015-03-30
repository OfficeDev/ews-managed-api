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