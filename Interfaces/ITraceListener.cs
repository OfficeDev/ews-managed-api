// ---------------------------------------------------------------------------
// <copyright file="ITraceListener.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ITraceListener interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;

    /// <summary>
    /// ITraceListener handles message tracing.
    /// </summary>
    public interface ITraceListener
    {
        /// <summary>
        /// Handles a trace message
        /// </summary>
        /// <param name="traceType">Type of trace message.</param>
        /// <param name="traceMessage">The trace message.</param>
        void Trace(string traceType, string traceMessage);
    }
}
