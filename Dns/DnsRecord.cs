// ---------------------------------------------------------------------------
// <copyright file="DnsRecord.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

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
