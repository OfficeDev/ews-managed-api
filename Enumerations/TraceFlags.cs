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
