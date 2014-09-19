// ---------------------------------------------------------------------------
// <copyright file="DnsSrvRecord.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//---------------------------------------------------------------------
// <summary>Defines the DnsSrvRecord class.</summary>
//---------------------------------------------------------------------
namespace Microsoft.Exchange.WebServices.Dns
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Represents a DNS SRV Record.
    /// </summary>
    internal class DnsSrvRecord : DnsRecord
    {
        /// <summary>The string representing the target host</summary>
        private string target;

        /// <summary>priority of the target host specified in the owner name.</summary>
        private int priority;

        /// <summary>weight of the target host</summary>
        private int weight;

        /// <summary>port used on the target for the service.</summary>
        private int port;

        /// <summary>
        /// Initializes a new instance of the DnsSrvRecord class.
        /// </summary>
        /// <param name="header">Dns dnsRecord header</param>
        /// <param name="dataPointer">Pointer to the data portion of the dnsRecord</param>
        internal override void Load(DnsRecordHeader header, IntPtr dataPointer)
        {
            base.Load(header, dataPointer);

            Win32DnsSrvRecord record = (Win32DnsSrvRecord)Marshal.PtrToStructure(dataPointer, typeof(Win32DnsSrvRecord));
            this.target = record.NameTarget;
            this.priority = record.Priority;
            this.weight = record.Weight;
            this.port = record.Port;
        }

        /// <summary>
        /// Gets the matching type of DNS dnsRecord.
        /// </summary>
        /// <value>The type of the dnsRecord.</value>
        internal override DnsRecordType RecordType
        {
            get { return DnsRecordType.SRV; }
        }

        /// <summary>
        /// Get the name target field of the DNS dnsRecord.
        /// </summary>
        internal string NameTarget
        {
            get { return this.target; }
        }

        /// <summary>
        /// Gwet the priority field of this DNS SRV Record.
        /// </summary>
        internal int Priority
        {
            get { return this.priority; }
        }

        /// <summary>
        /// Get the weight field of this DNS SRV Record.
        /// </summary>
        internal int Weight
        {
            get { return this.weight; }
        }

        /// <summary>
        /// Gets the port field of the DNS SRV dnsRecord.
        /// </summary>
        internal int Port
        {
            get { return this.port; }
        }

        /// <summary>
        ///  Win32DnsSrvRecord - native format SRV dnsRecord returned by DNS API
        /// </summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct Win32DnsSrvRecord
        {
            /// <summary>Represents the common DNS record header.</summary>
            public DnsRecordHeader Header;

            /// <summary>Represents the target host.</summary>
            public string NameTarget;

            /// <summary>Priority of the target host specified in the owner name. Lower numbers imply higher priority.</summary>
            public UInt16 Priority;

            /// <summary>
            /// Weight of the target host. Useful when selecting among hosts with the same priority. 
            /// The chances of using this host should be proportional to its weight
            /// </summary>
            public UInt16 Weight;

            /// <summary>Port used on the target host for the service.</summary>
            public UInt16 Port;

            /// <summary>Reserved. Used to keep pointers DWORD aligned.</summary>
            public UInt16 Pad; // keep ptrs ulong aligned
        }
    }
}
