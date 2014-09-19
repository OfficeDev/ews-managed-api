// ---------------------------------------------------------------------------
// <copyright file="DnsRecordType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//---------------------------------------------------------------------
// <summary>Defines the DnsRecordType enumeration.</summary>
//---------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Dns
{
    /// <summary>
    ///  DNS record types.
    /// </summary>
    internal enum DnsRecordType : ushort
    {
        /// <summary>
        ///  RFC 1034/1035 Address Record
        /// </summary>        
        A = 0x0001,

        /// <summary>
        /// Canonical Name Record
        /// </summary>
        CNAME = 0x0005,

        /// <summary>
        /// Start of Authority Record
        /// </summary>
        SOA = 0x0006,
        
        /// <summary>
        /// Pointer Record
        /// </summary>
        PTR = 0x000c,

        /// <summary>
        ///  Mail Exchange Record
        /// </summary>
        MX = 0x000f,

        /// <summary>
        /// Text Record
        /// </summary>
        TXT = 0x0010,

        /// <summary>
        ///  RFC 1886 (IPv6 Address)
        /// </summary>
        AAAA = 0x001c,

        /// <summary>
        /// Service location - RFC 2052
        /// </summary>
        SRV = 0x0021,
    }
}
