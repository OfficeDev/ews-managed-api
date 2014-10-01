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

//---------------------------------------------------------------------
// <summary>Defines the DnsClient class.</summary>
//---------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Dns
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Runtime.InteropServices;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// DNS Query client.
    /// </summary>
    internal class DnsClient
    {
        /// <summary>
        /// Win32 successful operation.</summary>
        private const int Win32Success = 0;

        /// <summary>
        /// Map type of DnsRecord to DnsRecordType.
        /// </summary>
        private static LazyMember<Dictionary<Type, DnsRecordType>> typeToDnsTypeMap = new LazyMember<Dictionary<Type, DnsRecordType>>(
            delegate()
            {
                Dictionary<Type, DnsRecordType> result = new Dictionary<Type, DnsRecordType>();
                result.Add(typeof(DnsSrvRecord), DnsRecordType.SRV);
                return result;
            });

        /// <summary>
        /// Perform DNS Query.
        /// </summary>
        /// <typeparam name="T">DnsRecord type.</typeparam>
        /// <param name="domain">The domain.</param>
        /// <param name="dnsServerAddress">IPAddress of DNS server to use (may be null).</param>
        /// <returns>The DNS record list (never null but may be empty).</returns>
        internal static List<T> DnsQuery<T>(string domain, IPAddress dnsServerAddress) where T : DnsRecord, new()
        {
            List<T> dnsRecordList = new List<T>();

            // Each strongly-typed DnsRecord type maps to a DnsRecordType enum.
            DnsRecordType dnsRecordTypeToQuery = typeToDnsTypeMap.Member[typeof(T)];

            // queryResultsPtr will point to unmanaged heap memoery if DnsQuery succeeds.
            IntPtr queryResultsPtr = IntPtr.Zero;

            try
            {
                // Perform DNS query. If successful, construct a list of results.
                int errorCode = DnsNativeMethods.DnsQuery(
                    domain,
                    dnsServerAddress,
                    dnsRecordTypeToQuery,
                    ref queryResultsPtr);

                if (errorCode == Win32Success)
                {
                    DnsRecordHeader dnsRecordHeader;

                    // Interate through linked list of query result records.
                    for (IntPtr recordPtr = queryResultsPtr; !recordPtr.Equals(IntPtr.Zero); recordPtr = dnsRecordHeader.NextRecord)
                    {
                        dnsRecordHeader = (DnsRecordHeader)Marshal.PtrToStructure(recordPtr, typeof(DnsRecordHeader));

                        T dnsRecord = new T();
                        if (dnsRecordHeader.RecordType == dnsRecord.RecordType)
                        {
                            dnsRecord.Load(dnsRecordHeader, recordPtr);
                            dnsRecordList.Add(dnsRecord);
                        }
                    }
                }
                else 
                {
                  throw new DnsException(errorCode);
                }
            }
            finally
            {
                if (queryResultsPtr != IntPtr.Zero)
                {
                    // DnsQuery allocated unmanaged heap, free it now.
                    DnsNativeMethods.FreeDnsQueryResults(queryResultsPtr);
                }
            }

            return dnsRecordList;
        }
    }
}
