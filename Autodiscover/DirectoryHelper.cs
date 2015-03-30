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

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.DirectoryServices;
    using System.DirectoryServices.ActiveDirectory;
    using System.Runtime.InteropServices;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represents a set of helper methods for using Active Directory services.
    /// </summary>
    internal class DirectoryHelper
    {
        #region Static members

        /// <summary>
        /// Maximum number of SCP hops in an SCP host lookup call.
        /// </summary>
        private const int AutodiscoverMaxScpHops = 10;

        /// <summary>
        /// GUID for SCP URL keyword
        /// </summary>
        private const string ScpUrlGuidString = @"77378F46-2C66-4aa9-A6A6-3E7A48B19596";

        /// <summary>
        /// GUID for SCP pointer keyword
        /// </summary>
        private const string ScpPtrGuidString = @"67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68";

        /// <summary>
        /// Filter string to find SCP Ptrs and Urls.
        /// </summary>
        private const string ScpFilterString = "(&(objectClass=serviceConnectionPoint)(|(keywords=" + ScpPtrGuidString + ")(keywords=" + ScpUrlGuidString + ")))";

        #endregion

        #region Private members

        private ExchangeServiceBase service;

        #endregion

        /// <summary>
        /// Gets the SCP URL list for domain.
        /// </summary>
        /// <param name="domainName">Name of the domain.</param>
        /// <returns>List of Autodiscover URLs</returns>
        public List<string> GetAutodiscoverScpUrlsForDomain(string domainName)
        {
            int maxHops = AutodiscoverMaxScpHops;
            List<string> scpUrlList;

            try
            {
                scpUrlList = this.GetScpUrlList(domainName, null, ref maxHops);
            }
            catch (InvalidOperationException e)
            {
                this.TraceMessage(
                    string.Format("LDAP call failed, exception: {0}", e.ToString()));
                scpUrlList = new List<string>();
            }
            catch (NotSupportedException e)
            {
                this.TraceMessage(
                    string.Format("LDAP call failed, exception: {0}", e.ToString()));
                scpUrlList = new List<string>();
            }
            catch (COMException e)
            {
                this.TraceMessage(
                    string.Format("LDAP call failed, exception: {0}", e.ToString()));
                scpUrlList = new List<string>();
            }

            return scpUrlList;
        }

        /// <summary>
        /// Search Active Directory for any related SCP URLs for a given domain name.
        /// </summary>
        /// <param name="domainName">Domain name to search for SCP information</param>
        /// <param name="ldapPath">LDAP path to start the search</param>
        /// <param name="maxHops">The number of remaining allowed hops</param>
        private List<string> GetScpUrlList(
           string domainName,
           string ldapPath,
           ref int maxHops)
        {
            if (maxHops <= 0)
            {
                throw new ServiceLocalException(Strings.MaxScpHopsExceeded);
            }

            maxHops--;

            this.TraceMessage(
                string.Format("Starting SCP lookup for domainName='{0}', root path='{1}'", domainName, ldapPath));

            string scpUrl = null;
            string fallBackLdapPath = null;
            string rootDsePath = null;
            string configPath = null;

            // The list of SCP URLs.
            List<string> scpUrlList = new List<string>();

            // Get the LDAP root path.
            rootDsePath = (ldapPath == null) ? "LDAP://RootDSE" : ldapPath + "/RootDSE";

            // Get the root directory entry.
            using (DirectoryEntry rootDseEntry = new DirectoryEntry(rootDsePath))
            {
                // Get the configuration path.
                configPath = rootDseEntry.Properties["configurationNamingContext"].Value as string;
            }

            // The container for SCP pointers and URLs objects from Active Directory
            SearchResultCollection scpDirEntries = null;

            try
            {
                // Get the configuration entry path.
                using (DirectoryEntry configEntry = new DirectoryEntry("LDAP://" + configPath))
                {
                    // Use the configuration entry path to create a query.
                    using (DirectorySearcher configSearcher = new DirectorySearcher(configEntry))
                    {
                        // Filter for Autodiscover SCP URLs and SCP pointers.
                        configSearcher.Filter = ScpFilterString;

                        // Identify properties to retrieve using the the query.
                        configSearcher.PropertiesToLoad.Add("keywords");
                        configSearcher.PropertiesToLoad.Add("serviceBindingInformation");

                        this.TraceMessage(
                            string.Format("Searching for SCP entries in {0}", configEntry.Path));

                        // Query Active Directory for SCP entries.
                        scpDirEntries = configSearcher.FindAll();
                    }
                }

                // Identify the domain to match.
                string domainMatch = "Domain=" + domainName;

                // Contains a pointer to the LDAP path of a SCP object.
                string scpPtrLdapPath;

                this.TraceMessage(
                    string.Format("Scanning for SCP pointers {0}", domainMatch));

                foreach (SearchResult scpDirEntry in scpDirEntries)
                {
                    ResultPropertyValueCollection entryKeywords = scpDirEntry.Properties["keywords"];

                    // Identify SCP pointers.
                    if (entryKeywords.CaseInsensitiveContains(ScpPtrGuidString))
                    {
                        // Get the LDAP path to SCP pointer.
                        scpPtrLdapPath = scpDirEntry.Properties["serviceBindingInformation"][0] as string;

                        // If the SCP pointer matches the user's domain, then restart search from that point.
                        if (entryKeywords.CaseInsensitiveContains(domainMatch))
                        {
                            // Stop the current search, start another from a new location.
                            this.TraceMessage(
                                string.Format(
                                    "SCP pointer for '{0}' is found in '{1}', restarting seach in '{2}'",
                                    domainMatch,
                                    scpDirEntry.Path,
                                    scpPtrLdapPath));

                            return this.GetScpUrlList(domainName, scpPtrLdapPath, ref maxHops);
                        }
                        else
                        {
                            // Save the first SCP pointer ldapPath for a later call if a SCP URL is not found.
                            // Directory entries with a SCP pointer should have only one keyword=ScpPtrGuidString.
                            if ((entryKeywords.Count == 1) && string.IsNullOrEmpty(fallBackLdapPath))
                            {
                                fallBackLdapPath = scpPtrLdapPath;
                                this.TraceMessage(
                                    string.Format(
                                        "Fallback SCP pointer='{0}' for '{1}' is found in '{2}' and saved.",
                                        fallBackLdapPath,
                                        domainMatch,
                                        scpDirEntry.Path));
                            }
                        }
                    }
                }

                this.TraceMessage(
                    string.Format("No SCP pointers found for '{0}' in configPath='{1}'", domainMatch, configPath));

                // Get the computer's current site.
                string computerSiteName = this.GetSiteName();

                if (!string.IsNullOrEmpty(computerSiteName))
                {
                    // Search for SCP entries.
                    string sitePrefix = "Site=";
                    string siteMatch = sitePrefix + computerSiteName;
                    List<string> scpListNoSiteMatch = new List<string>();

                    this.TraceMessage(
                        string.Format("Scanning for SCP urls for the current computer {0}", siteMatch));

                    foreach (SearchResult scpDirEntry in scpDirEntries)
                    {
                        ResultPropertyValueCollection entryKeywords = scpDirEntry.Properties["keywords"];

                        // Identify SCP URLs.
                        if (entryKeywords.CaseInsensitiveContains(ScpUrlGuidString) && scpDirEntry.Properties["serviceBindingInformation"].Count > 0)
                        {
                            // Get the SCP URL.
                            scpUrl = scpDirEntry.Properties["serviceBindingInformation"][0] as string;

                            // If the SCP URL matches the exact ComputerSiteName.
                            if (entryKeywords.CaseInsensitiveContains(siteMatch))
                            {
                                // Priority 1 SCP URL. Add SCP URL to the list if it's not already there.
                                if (!scpUrlList.CaseInsensitiveContains(scpUrl))
                                {
                                    this.TraceMessage(
                                        string.Format(
                                            "Adding (prio 1) '{0}' for the '{1}' from '{2}' to the top of the list (exact match)",
                                            scpUrl,
                                            siteMatch,
                                            scpDirEntry.Path));

                                    scpUrlList.Add(scpUrl);
                                }
                            }

                            // No match between the SCP URL and the ComputerSiteName
                            else
                            {
                                bool hasSiteKeyword = false;

                                // Check if SCP URL entry has any keyword starting with "Site=" 
                                foreach (string keyword in entryKeywords)
                                {
                                    hasSiteKeyword |= keyword.StartsWith(sitePrefix, StringComparison.OrdinalIgnoreCase);
                                }

                                // Add SCP URL to the scpListNoSiteMatch list if it's not already there.
                                if (!scpListNoSiteMatch.CaseInsensitiveContains(scpUrl))
                                {
                                    // Priority 2 SCP URL. SCP entry doesn't have any "Site=<otherSite>" keywords, insert at the top of list.
                                    if (!hasSiteKeyword)
                                    {
                                        this.TraceMessage(
                                            string.Format(
                                                "Adding (prio 2) '{0}' from '{1}' to the middle of the list (wildcard)",
                                                scpUrl,
                                                scpDirEntry.Path));

                                        scpListNoSiteMatch.Insert(0, scpUrl);
                                    }

                                    // Priority 3 SCP URL. SCP entry has at least one "Site=<otherSite>" keyword, add to the end of list.
                                    else
                                    {
                                        this.TraceMessage(
                                            string.Format(
                                                "Adding (prio 3) '{0}' from '{1}' to the end of the list (site mismatch)",
                                                scpUrl,
                                                scpDirEntry.Path));

                                        scpListNoSiteMatch.Add(scpUrl);
                                    }
                                }
                            }
                        }
                    }

                    // Append SCP URLs to the list. List contains:
                    // Priority 1 URLs -- URLs with an exact match, "Site=<machineSite>"
                    // Priority 2 URLs -- URLs without a match, no any "Site=<anySite>" in the entry
                    // Priority 3 URLs -- URLs without a match, "Site=<nonMachineSite>"
                    if (scpListNoSiteMatch.Count > 0)
                    {
                        foreach (string url in scpListNoSiteMatch)
                        {
                            if (!scpUrlList.CaseInsensitiveContains(url))
                            {
                                scpUrlList.Add(url);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (scpDirEntries != null)
                {
                    scpDirEntries.Dispose();
                }
            }

            // If no entries found, try fallBackLdapPath if it's non-empty.
            if (scpUrlList.Count == 0)
            {
                if (!string.IsNullOrEmpty(fallBackLdapPath))
                {
                    this.TraceMessage(
                        string.Format(
                        "Restarting search for domain '{0}' in SCP fallback pointer '{1}'",
                        domainName,
                        fallBackLdapPath));

                    return this.GetScpUrlList(domainName, fallBackLdapPath, ref maxHops);
                }
            }

            // Return the list with 0 or more SCP URLs.
            return scpUrlList;
        }

        /// <summary>
        /// Get the local site name.
        /// </summary>
        /// <returns>Name of the local site.</returns>
        private string GetSiteName()
        {
            try
            {
                using (ActiveDirectorySite site = ActiveDirectorySite.GetComputerSite())
                {
                    return site.Name;
                }
            }
            catch (ActiveDirectoryObjectNotFoundException)  // object not found in directory store
            {
                return null;
            }
            catch (ActiveDirectoryOperationException)       // underlying directory operation failed
            {
                return null;
            }
            catch (ActiveDirectoryServerDownException)      // server unavailable
            {
                return null;
            }
        }

        /// <summary>
        /// Traces message.
        /// </summary>
        /// <param name="message">The message.</param>
        private void TraceMessage(string message)
        {
            this.Service.TraceMessage(TraceFlags.AutodiscoverConfiguration, message);
        }

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DirectoryHelper"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        public DirectoryHelper(ExchangeServiceBase service)
        {
            this.service = service;
        }
        #endregion

        #region Properties

        internal ExchangeServiceBase Service
        {
            get { return this.service; }
        }

        #endregion
    }
}