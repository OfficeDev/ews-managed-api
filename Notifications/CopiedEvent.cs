// ---------------------------------------------------------------------------
// <copyright file="CopiedEvent.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the CopiedEvent class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Represents an event indicating that an item or folder was moved or copied.
    /// </summary>
    public sealed class CopiedEvent : MovedCopiedEvent
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CopiedEvent"/> class.
        /// </summary>
        internal CopiedEvent()
            : base()
        {
        }
    }
}
