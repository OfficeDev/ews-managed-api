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

//------------------------------------------------------------------------------
// <summary>Defines the AutodiscoverDnsClient class.</summary>
//------------------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    using System;
    using System.Collections.Generic;
    using System.Security;
    using Microsoft.Exchange.WebServices.Data;
    using Microsoft.Exchange.WebServices.Dns;

    /// <summary>
    /// Class that reads AutoDiscover configuration information from DNS.
    /// </summary>
    internal class AutodiscoverDnsClient
    {
        #region Constants
        /// <summary>
        /// SRV DNS prefix to lookup.
        /// </summary>
        private const string AutoDiscoverSrvPrefix = "_autodiscover._tcp.";

        /// <summary>
        /// We are only interested in records that use SSL.
        /// </summary>
        private const int SslPort = 443;
        #endregion

        #region Static fields
        /// <summary>
        /// Random selector in the case of ties.
        /// </summary>
        private static Random randomTieBreakerSelector = new Random();
        #endregion

        #region Instance fields
        /// <summary>
        /// AutodiscoverService using this DNS reader.
        /// </summary>
        private AutodiscoverService service;

        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverDnsClient"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal AutodiscoverDnsClient(AutodiscoverService service)
        {
            this.service = service;
        }

        #endregion

        #region Instance methods

        /// <summary>
        /// Finds the Autodiscover host from DNS SRV records.
        /// </summary>
        /// <remarks>
        /// If the domain to lookup is "contoso.com", Autodiscover will use DnsQuery on SRV records 
        /// for "_autodiscover._tcp.contoso.com". If the query is successful it will return a target 
        /// domain (e.g. "mail.contoso.com") which will be tried as an Autodiscover endpoint.
        /// </remarks>
        /// <param name="domain">The domain.</param>
        /// <returns>Autodiscover hostname (will be null if lookup failed).</returns>
        internal string FindAutodiscoverHostFromSrv(string domain)
        {
            string domainToMatch = AutoDiscoverSrvPrefix + domain;

            DnsSrvRecord dnsSrvRecord = this.FindBestMatchingSrvRecord(domainToMatch);

            if ((dnsSrvRecord == null) || string.IsNullOrEmpty(dnsSrvRecord.NameTarget))
            {
                this.service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    "No appropriate SRV record was found.");
                return null;
            }
            else
            {
                this.service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("DNS query for SRV record for domain {0} found {1}", domain, dnsSrvRecord.NameTarget));

                return dnsSrvRecord.NameTarget;
            }
        }

        /// <summary>
        /// Finds the best matching SRV record.
        /// </summary>
        /// <param name="domain">The domain.</param>
        /// <returns>DnsSrvRecord(will be null if lookup failed).</returns>
        private DnsSrvRecord FindBestMatchingSrvRecord(string domain)
        {
            List<DnsSrvRecord> dnsSrvRecordList;
            try
            {
                // Make DnsQuery call to get collection of SRV records.
                dnsSrvRecordList = DnsClient.DnsQuery<DnsSrvRecord>(domain, this.service.DnsServerAddress);
            }
            catch (DnsException ex)
            {
                string dnsExcMessage = string.Format(
                        "DnsQuery returned error error '{0}' error code 0x{1:X8}.",
                        ex.Message,
                        ex.NativeErrorCode);
                this.service.TraceMessage(TraceFlags.AutodiscoverConfiguration, dnsExcMessage);
                return null;
            }
            catch (SecurityException ex)
            {
                // In restricted environments, we may not be allowed to call unmanaged code.
                this.service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("DnsQuery cannot be called. Security error: {0}.", ex.Message));
                return null;
            }

            this.service.TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("{0} SRV records were returned.", dnsSrvRecordList.Count));

            // If multiple records were returned, they will be returned sorted by priority 
            // (and weight) order. Need to find the index of the first record that supports SSL.
            int priority = int.MinValue;
            int weight = int.MinValue;
            bool recordFound = false;
            foreach (DnsSrvRecord dnsSrvRecord in dnsSrvRecordList)
            {
                if (dnsSrvRecord.Port == SslPort)
                {
                    priority = dnsSrvRecord.Priority;
                    weight = dnsSrvRecord.Weight;
                    recordFound = true;
                    break;
                }
            }

            // Records were returned but nothing matched our criteria.
            if (!recordFound)
            {
                this.service.TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    "No appropriate SRV records were found.");

                return null;
            }

            // Collect all records with the same (highest) priority.
            // (Aren't lambda expressions cool? ;-)
            List<DnsSrvRecord> bestDnsSrvRecordList = dnsSrvRecordList.FindAll(
                record => (record.Port == SslPort) && (record.Priority == priority) && (record.Weight == weight));

            // The list must contain at least one matching record since we found one earlier.
            EwsUtilities.Assert(
                dnsSrvRecordList.Count > 0,
                "AutodiscoverDnsClient.FindBestMatchingSrvRecord",
                "At least one DNS SRV record must match the criteria.");

            // If we have multiple records with the same priority and weight, randomly pick one.
            int recordIndex = (bestDnsSrvRecordList.Count > 1)
                ? randomTieBreakerSelector.Next(bestDnsSrvRecordList.Count) 
                : 0;

            DnsSrvRecord bestDnsSrvRecord = bestDnsSrvRecordList[recordIndex];

            string traceMessage = string.Format(
                 "Returning SRV record {0} of {1} records. Target: {2}, Priority: {3}, Weight: {4}",
                 recordIndex,
                 dnsSrvRecordList.Count,
                 bestDnsSrvRecord.NameTarget,
                 bestDnsSrvRecord.Priority,
                 bestDnsSrvRecord.Weight);
             this.service.TraceMessage( TraceFlags.AutodiscoverConfiguration, traceMessage);

            return bestDnsSrvRecord;
        }
        #endregion
    }
}