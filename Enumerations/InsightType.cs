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

using System.Xml.Serialization;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Defines the type of an InsightType object.
    /// </summary>
    public enum InsightType
    {
        /// <summary>
        /// The InsightType represents the full name.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        FullName,

        /// <summary>
        /// The InsightType represents the first name.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        FirstName,

        /// <summary>
        /// The InsightType represents the last name.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        LastName,

        /// <summary>
        /// The InsightType represents the phone number.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        PhoneNumber,

        /// <summary>
        /// The InsightType represents the SMS number.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        SMSNumber,

        /// <summary>
        /// The InsightType represents the email address
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        EmailAddress,

        /// <summary>
        /// The InsightType represents the facebook profile link.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        FacebookProfileLink,

        /// <summary>
        /// The InsightType represents the linkedin profile link.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        LinkedInProfileLink,

        /// <summary>
        /// The InsightType represents the provious job.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        PreviousJob,

        /// <summary>
        /// The InsightType represents the education history.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        EducationHistory,

        /// <summary>
        /// The InsightType represents the skills.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Skills,

        /// <summary>
        /// The InsightType represents the professional biography.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        ProfessionalBiography,

        /// <summary>
        /// The InsightType represents the management chain.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        ManagementChain,

        /// <summary>
        /// The InsightType represents the direct reports.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        DirectReports,

        /// <summary>
        /// The InsightType represents the peers.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Peers,

        /// <summary>
        /// The InsightType represents the team size.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        TeamSize,

        /// <summary>
        /// The InsightType represents the current job.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        CurrentJob,

        /// <summary>
        /// The InsightType represents the birthday.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Birthday,

        /// <summary>
        /// The InsightType represents the home town.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Hometown,

        /// <summary>
        /// The InsightType represents the current location.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        CurrentLocation,

        /// <summary>
        /// The InsightType represents the user profile picture.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        UserProfilePicture,

        /// <summary>
        /// The InsightType represents the Delve doc.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        DelveDoc,

        /// <summary>
        /// The InsightType represents the company profile.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        CompanyProfile,

        /// <summary>
        /// The InsightType represents the Office.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Office,
        
        /// <summary>
        /// The InsightType represents the Headline.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Headline,

        /// <summary>
        /// The InsightType represents the mutual connection.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        MutualConnections,

        /// <summary>
        /// The InsightType represents the title.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        Title,

        /// <summary>
        /// The InsightType represents the mutual manager.
        /// </summary>
        [RequiredServerVersion(ExchangeVersion.Exchange2015)]
        MutualManager,
    }
}