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