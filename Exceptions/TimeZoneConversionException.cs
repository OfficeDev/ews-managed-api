// ---------------------------------------------------------------------------
// <copyright file="TimeZoneConversionException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ServiceLocalException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when a date and time cannot be converted from one time zone
    /// to another.
    /// </summary>
    [Serializable]
    public class TimeZoneConversionException : ServiceLocalException
    {
        /// <summary>
        /// ServiceLocalException Constructor.
        /// </summary>
        public TimeZoneConversionException()
            : base()
        {
        }

        /// <summary>
        /// ServiceLocalException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        public TimeZoneConversionException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// ServiceLocalException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="innerException">Inner exception.</param>
        public TimeZoneConversionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}