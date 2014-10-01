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
// <summary>Defines the DnsRecord class.</summary>
//---------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Dns
{
    using System;

    /// <summary>
    /// Represents a DNS Record.
    /// </summary>
    internal abstract class DnsRecord
    {
        /// <summary>
        /// Name field of this DNS Record.
        /// </summary>
        private string name;

        /// <summary>
        /// The suggested time for this dnsRecord to be valid.
        /// </summary>
        private uint timeToLive;

        /// <summary>
        /// Loads the DNS dnsRecord.
        /// </summary>
        /// <param name="header">The header.</param>
        /// <param name="dataPointer">The data pointer.</param>
        internal virtual void Load(DnsRecordHeader header, IntPtr dataPointer)
        {
            this.name = header.Name;
            this.timeToLive = Math.Max(1, header.Ttl);
        }

        /// <summary>
        /// Gets the type of the DnsRecord.
        /// </summary>
        /// <value>The type of the DnsRecord.</value>
        internal abstract DnsRecordType RecordType
        {
            get;
        }

        /// <summary>
        /// Name property
        /// </summary>
        public string Name
        {
            get { return this.name; }
        }

        /// <summary>
        /// The suggested duration that this dnsRecord is valid
        /// </summary>
        public TimeSpan TimeToLive
        {
            get { return TimeSpan.FromSeconds((double)this.timeToLive); }
        }
    }
}
