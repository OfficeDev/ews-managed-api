// ---------------------------------------------------------------------------
// <copyright file="DnsRecordHeader.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//---------------------------------------------------------------------
// <summary>Defines the DnsRecordHeader class.</summary>
//---------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Dns
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    ///  Represents the native format of a DNS record returned by the Win32 DNS API
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct DnsRecordHeader
    {
        /// <summary>
        /// Pointer to the next DNS dnsRecord.
        /// </summary>
        internal IntPtr NextRecord;

        /// <summary>
        /// Domain name of the dnsRecord set to be updated.
        /// </summary>
        internal string Name;

        /// <summary>The type of the current dnsRecord.</summary>
        internal DnsRecordType RecordType;

        /// <summary>Length of the data, in bytes. </summary>
        internal UInt16 DataLength;

        /// <summary>
        /// Flags used in the structure, in the form of a bit-wise DWORD.
        /// </summary>
        internal UInt32 Flags;

        /// <summary>
        /// Time to live, in seconds
        /// </summary>
        internal UInt32 Ttl;

        /// <summary>
        /// Reserved for future use.
        /// </summary>
        internal UInt32 Reserved;
    }
}
