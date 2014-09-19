// ---------------------------------------------------------------------------
// <copyright file="JsonDeserializationNotImplementedException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Exception thrown for a method that does not support Json deserialization
    /// </summary>
    [Serializable]
    internal class JsonDeserializationNotImplementedException : ServiceLocalException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="JsonDeserializationNotImplementedException"/> class.
        /// </summary>
        internal JsonDeserializationNotImplementedException() :
            base(Strings.JsonDeserializationNotImplemented)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="JsonDeserializationNotImplementedException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        internal JsonDeserializationNotImplementedException(string message) :
            base(message)
        {
        }
    }
}
