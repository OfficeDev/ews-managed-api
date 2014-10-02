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
// <summary>Defines the DnsNativeMethods class.</summary>
//---------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Dns
{
    using System;
    using System.Diagnostics;
    using System.Diagnostics.CodeAnalysis;
    using System.Net;
    using System.Net.Sockets;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Class that defined native Win32 DNS API methods
    /// </summary>
    [ComVisible(false)]
    internal static class DnsNativeMethods
    {
        /// <summary>
        /// The Win32 dll from which to load DNS APIs.
        /// </summary>
        /// <remarks>
        /// DNSAPI.DLL has been part of the Win32 API since Win2K. Don't need to verify that the DLL exists.
        /// </remarks>
        private const string DNSAPI = "dnsapi.dll";

        /// <summary>
        /// Win32 memory free type enumeration.
        /// </summary>
        /// <remarks>Win32 defines other values for this enum but we don't uses them.</remarks>
        private enum FreeType
        {
            /// <summary>
            /// The data freed is a Resource Record list, and includes subfields of the DNS_RECORD
            /// structure. Resources freed include structures returned by the DnsQuery and DnsRecordSetCopyEx functions.
            /// </summary>
            RecordList = 1,
        }

        /// <summary>
        /// DNS Query options.
        /// </summary>
        /// <remarks>Win32 defines other values for this enum but we don't uses them.</remarks>
        private enum DnsQueryOptions
        {
            /// <summary>
            /// Default option.
            /// </summary>
            DNS_QUERY_STANDARD = 0,
        }

        /// <summary>
        /// Represents the native format of a DNS record returned by the Win32 DNS API
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DnsServerList
        {
            public Int32 AddressCount;
            public Int32 ServerAddress;
        }

        /// <summary>
        /// Call Win32 DNS API DnsQuery.
        /// </summary>
        /// <param name="pszName">Host name.</param>
        /// <param name="wType">DNS Record type.</param>
        /// <param name="options">DNS Query options.</param>
        /// <param name="aipServers">Array of DNS server IP addresses.</param>
        /// <param name="ppQueryResults">Query results.</param>
        /// <param name="pReserved">Reserved argument.</param>
        /// <returns>WIN32 status code</returns>
        /// <remarks>For aipServers, DnqQuery expects either null or an array of one IPv4 address.</remarks>
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage", Justification = "Managed API")]
        [DllImport(DNSAPI, EntryPoint = "DnsQuery_W", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
        private static extern int DnsQuery(
            [In] string pszName,
            DnsRecordType wType,
            DnsQueryOptions options,
            IntPtr aipServers,
            ref IntPtr ppQueryResults,
            int pReserved);

        /// <summary>
        /// Call Win32 DNS API DnsRecordListFree.
        /// </summary>
        /// <param name="ptrRecords">DNS records pointer</param>
        /// <param name="freeType">Record List Free type</param>
        [SuppressMessage("Microsoft.Security", "CA2118:ReviewSuppressUnmanagedCodeSecurityUsage", Justification = "Managed API")]
        [DllImport(DNSAPI, EntryPoint = "DnsRecordListFree", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode)]
        private static extern void DnsRecordListFree([In] IntPtr ptrRecords, [In] FreeType freeType);

        /// <summary>
        /// Allocate the DNS server list.
        /// </summary>
        /// <param name="dnsServerAddress">The DNS server address (may be null).</param>
        /// <returns>Pointer to DNS server list (may be IntPtr.Zero).</returns>
        private static IntPtr AllocDnsServerList(IPAddress dnsServerAddress)
        {
            IntPtr pServerList = IntPtr.Zero;

            // Build DNS server list arg if DNS server address was passed in.
            // Note: DnsQuery only supports a single IP address and it has to 
            // be an IPv4 address.
            if (dnsServerAddress != null)
            {
                Debug.Assert(
                    dnsServerAddress.AddressFamily == AddressFamily.InterNetwork,
                    "DnsNativeMethods",
                    "Only Ipv4 DNS server addresses are supported by DnsQuery");

                Int32 serverAddress = BitConverter.ToInt32(dnsServerAddress.GetAddressBytes(), 0);
                DnsServerList serverList;
                serverList.AddressCount = 1;
                serverList.ServerAddress = serverAddress;

                pServerList = Marshal.AllocHGlobal(Marshal.SizeOf(serverList));
                Marshal.StructureToPtr(serverList, pServerList, false);
            }
            return pServerList;
        }

        /// <summary>
        /// Wrapper method to perform DNS Query.
        /// </summary>
        /// <remarks>Makes DnsQuery a little more palatable.</remarks>
        /// <param name="domain">The domain.</param>
        /// <param name="dnsServerAddress">IPAddress of DNS server (may be null) </param>
        /// <param name="recordType">Type of DNS dnsRecord.</param>
        /// <param name="ppQueryResults">Pointer to pointer to query results.</param>
        /// <returns>Win32 status code.</returns>
        internal static int DnsQuery(
            string domain,
            IPAddress dnsServerAddress,
            DnsRecordType recordType,
            ref IntPtr ppQueryResults)
        {
            Debug.Assert( !string.IsNullOrEmpty(domain), "domain cannot be null.");

            IntPtr pServerList = IntPtr.Zero;

            try
            {
                pServerList = DnsNativeMethods.AllocDnsServerList(dnsServerAddress);

                return DnsNativeMethods.DnsQuery(
                    domain,
                    recordType,
                    DnsQueryOptions.DNS_QUERY_STANDARD,
                    pServerList,
                    ref ppQueryResults,
                    0);
            }
            finally
            {
                // Note: if pServerList is IntPtr.Zero, FreeHGlobal does nothing.
                Marshal.FreeHGlobal(pServerList);
            }
        }

        /// <summary>
        /// Free results from DnsQuery call.
        /// </summary>
        /// <remarks>Makes DnsRecordListFree a little more palatable.</remarks>
        /// <param name="ptrRecords">Pointer to records.</param>
        internal static void FreeDnsQueryResults(IntPtr ptrRecords)
        {
            Debug.Assert( !ptrRecords.Equals(IntPtr.Zero), "ptrRecords cannot be null.");

            DnsNativeMethods.DnsRecordListFree( ptrRecords, FreeType.RecordList);
        }
    }
}
