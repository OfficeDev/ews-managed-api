// ---------------------------------------------------------------------------
// <copyright file="TraceFlags.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the TraceFlags enumeration.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// Defines flags to control tracing details.
    /// </summary>
    [Flags]
    public enum TraceFlags : long
    {
        /// <summary>
        /// No tracing.
        /// </summary>
        None = 0,

        /// <summary>
        /// Trace EWS request messages.
        /// </summary>
        EwsRequest = 1,

        /// <summary>
        /// Trace EWS response messages.
        /// </summary>
        EwsResponse = 2,

        /// <summary>
        /// Trace EWS response HTTP headers.
        /// </summary>
        EwsResponseHttpHeaders = 4,

        /// <summary>
        /// Trace Autodiscover request messages.
        /// </summary>
        AutodiscoverRequest = 8,

        /// <summary>
        /// Trace Autodiscover response messages.
        /// </summary>
        AutodiscoverResponse = 16,

        /// <summary>
        /// Trace Autodiscover response HTTP headers.
        /// </summary>
        AutodiscoverResponseHttpHeaders = 32,

        /// <summary>
        /// Trace Autodiscover configuration logic.
        /// </summary>
        AutodiscoverConfiguration = 64,

        /// <summary>
        /// Trace messages used in debugging the Exchange Web Services Managed API
        /// </summary>
        DebugMessage = 128,

        /// <summary>
        /// Trace EWS request HTTP headers.
        /// </summary>
        EwsRequestHttpHeaders = 256,

        /// <summary>
        /// Trace Autodiscover request HTTP headers.
        /// </summary>
        AutodiscoverRequestHttpHeaders = 512,

        /// <summary>
        /// All trace types enabled.
        /// </summary>
        All = Int64.MaxValue,
    }
}
