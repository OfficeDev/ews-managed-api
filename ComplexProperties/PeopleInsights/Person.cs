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
    /// <summary>
    /// Represents the Person.
    /// </summary>
    public class Person : ComplexProperty
    {
        /// <summary>
        /// Initializes a local instance of <see cref="Person"/>
        /// </summary>
        public Person()
            : base()
        {
        }

        /// <summary>
        /// Gets the EmailAddress.
        /// </summary>
        public string EmailAddress
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the collection of insights.
        /// </summary>
        public PersonInsightCollection Insights
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the display name.
        /// </summary>
        public string FullName
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the display name.
        /// </summary>
        public string DisplayName
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the given name.
        /// </summary>
        public string GivenName
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the surname.
        /// </summary>
        public string Surname
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the phone number.
        /// </summary>
        public string PhoneNumber 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the sms number.
        /// </summary>
        public string SMSNumber 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the facebook profile link.
        /// </summary>
        public string FacebookProfileLink 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the linkedin profile link.
        /// </summary>
        public string LinkedInProfileLink 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the list of skills.
        /// </summary>
        public SkillInsightValueCollection Skills
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the professional biography.
        /// </summary>
        public string ProfessionalBiography 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the management chain.
        /// </summary>
        public ProfileInsightValueCollection ManagementChain
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the list of direct reports.
        /// </summary>
        public ProfileInsightValueCollection DirectReports
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the list of peers.
        /// </summary>
        public ProfileInsightValueCollection Peers 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the team size.
        /// </summary>
        public string TeamSize 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the current job.
        /// </summary>
        public JobInsightValueCollection CurrentJob 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the birthday.
        /// </summary>
        public string Birthday 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the hometown.
        /// </summary>
        public string Hometown 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the current location.
        /// </summary>
        public string CurrentLocation 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the company profile.
        /// </summary>
        public CompanyInsightValueCollection CompanyProfile 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the office.
        /// </summary>
        public string Office 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the headline.
        /// </summary>
        public string Headline 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the list of mutual connections.
        /// </summary>
        public ProfileInsightValueCollection MutualConnections 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the title.
        /// </summary>
        public string Title 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the mutual manager.
        /// </summary>
        public ProfileInsightValue MutualManager 
        { 
            get; 
            internal set; 
        }

        /// <summary>
        /// Gets the alias.
        /// </summary>
        public string Alias
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the Department.
        /// </summary>
        public string Department
        {
            get;
            internal set;
        }

        /// <summary>
        /// Gets the user profile picture.
        /// </summary>
        public UserProfilePicture UserProfilePicture
        {
            get;
            internal set;
        }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
        {
            switch (reader.LocalName)
            {
                case XmlElementNames.EmailAddress:
                    this.EmailAddress = reader.ReadElementValue();
                    break;
                case XmlElementNames.FullName:
                    this.FullName = reader.ReadElementValue();
                    break;
                case XmlElementNames.DisplayName:
                    this.DisplayName = reader.ReadElementValue();
                    break;
                case XmlElementNames.GivenName:
                    this.GivenName = reader.ReadElementValue();
                    break;
                case XmlElementNames.Surname:
                    this.Surname = reader.ReadElementValue();
                    break;
                case XmlElementNames.PhoneNumber:
                    this.PhoneNumber = reader.ReadElementValue();
                    break;
                case XmlElementNames.SMSNumber:
                    this.SMSNumber = reader.ReadElementValue();
                    break;
                case XmlElementNames.FacebookProfileLink:
                    this.FacebookProfileLink = reader.ReadElementValue();
                    break;
                case XmlElementNames.LinkedInProfileLink:
                    this.LinkedInProfileLink = reader.ReadElementValue();
                    break;
                case XmlElementNames.ProfessionalBiography:
                    this.ProfessionalBiography = reader.ReadElementValue();
                    break;
                case XmlElementNames.TeamSize:
                    this.TeamSize = reader.ReadElementValue();
                    break;
                case XmlElementNames.Birthday:
                    this.Birthday = reader.ReadElementValue();
                    break;
                case XmlElementNames.Hometown:
                    this.Hometown = reader.ReadElementValue();
                    break;
                case XmlElementNames.CurrentLocation:
                    this.CurrentLocation = reader.ReadElementValue();
                    break;
                case XmlElementNames.Office:
                    this.Office = reader.ReadElementValue();
                    break;
                case XmlElementNames.Headline:
                    this.Headline = reader.ReadElementValue();
                    break;
                case XmlElementNames.Title:
                    this.Title = reader.ReadElementValue();
                    break;
                case XmlElementNames.Alias:
                    this.Alias = reader.ReadElementValue();
                    break;
                case XmlElementNames.Department:
                    this.Department = reader.ReadElementValue();
                    break;
                case XmlElementNames.MutualManager:
                    this.MutualManager = new ProfileInsightValue();
                    this.MutualManager.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.MutualManager);
                    break;
                case XmlElementNames.ManagementChain:
                    this.ManagementChain = new ProfileInsightValueCollection(XmlElementNames.Item);
                    this.ManagementChain.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ManagementChain);
                    break;
                case XmlElementNames.DirectReports:
                    this.DirectReports = new ProfileInsightValueCollection(XmlElementNames.Item);
                    this.DirectReports.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.DirectReports);
                    break;
                case XmlElementNames.Peers:
                    this.Peers = new ProfileInsightValueCollection(XmlElementNames.Item);
                    this.Peers.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Peers);
                    break;
                case XmlElementNames.MutualConnections:
                    this.MutualConnections = new ProfileInsightValueCollection(XmlElementNames.Item);
                    this.MutualConnections.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.MutualConnections);
                    break;
                case XmlElementNames.Skills:
                    this.Skills = new SkillInsightValueCollection(XmlElementNames.Item);
                    this.Skills.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Skills);
                    break;
                case XmlElementNames.CurrentJob:
                    this.CurrentJob = new JobInsightValueCollection(XmlElementNames.Item);
                    this.CurrentJob.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.CurrentJob);
                    break;
                case XmlElementNames.CompanyProfile:
                    this.CompanyProfile = new CompanyInsightValueCollection(XmlElementNames.Item);
                    this.CompanyProfile.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.CompanyProfile);
                    break;
                case XmlElementNames.Insights:
                    this.Insights = new PersonInsightCollection();
                    this.Insights.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.Insights);
                    break;
                case XmlElementNames.UserProfilePicture:
                    this.UserProfilePicture = new UserProfilePicture();
                    this.UserProfilePicture.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.UserProfilePicture);
                    break;
                default:
                    return base.TryReadElementFromXml(reader);
            }

            return true;
        }
    }
}