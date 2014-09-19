// ---------------------------------------------------------------------------
// <copyright file="EwsTraceListener.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the EwsTraceListener class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.IO;

    /// <summary>
    /// EwsTraceListener logs request/responses to a text writer.
    /// </summary>
    internal class EwsTraceListener : ITraceListener
    {
        private TextWriter writer;

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsTraceListener"/> class.
        /// Uses Console.Out as output.
        /// </summary>
        internal EwsTraceListener()
            : this(Console.Out)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsTraceListener"/> class.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal EwsTraceListener(TextWriter writer)
        {
            this.writer = writer;
        }

        /// <summary>
        /// Handles a trace message
        /// </summary>
        /// <param name="traceType">Type of trace message.</param>
        /// <param name="traceMessage">The trace message.</param>
        public void Trace(string traceType, string traceMessage)
        {
            this.writer.Write(traceMessage);
        }
    }
}
