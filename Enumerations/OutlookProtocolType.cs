// ---------------------------------------------------------------------------
// <copyright file="OutlookProtocolType.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the OutlookProtocolType enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Autodiscover
{
    /// <summary>
    /// Defines supported Outlook protocls.
    /// </summary>
    internal enum OutlookProtocolType
    {
        /// <summary>
        /// The Remote Procedure Call (RPC) protocol.
        /// </summary>
        Rpc,

        /// <summary>
        /// The Remote Procedure Call (RPC) over HTTP protocol.
        /// </summary>
        RpcOverHttp,

        /// <summary>
        /// The Web protocol.
        /// </summary>
        Web,

        /// <summary>
        /// The protocol is unknown.
        /// </summary>
        Unknown,
    }
}
