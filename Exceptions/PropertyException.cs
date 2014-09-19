// ---------------------------------------------------------------------------
// <copyright file="PropertyException.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the PropertyException class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Text;

    /// <summary>
    /// Represents an error that occurs when an operation on a property fails.
    /// </summary>
    [Serializable]
    public class PropertyException : ServiceLocalException
    {
        /// <summary>
        /// The name of the property that is at the origin of the exception.
        /// </summary>
        private string name;

        /// <summary>
        /// PropertyException constructor.
        /// </summary>
        /// <param name="name">The name of the property that is at the origin of the exception.</param>
        public PropertyException(string name)
            : base()
        {
            this.name = name;
        }

        /// <summary>
        /// PropertyException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="name">The name of the property that is at the origin of the exception.</param>
        public PropertyException(string message, string name)
            : base(message)
        {
            this.name = name;
        }

        /// <summary>
        /// PropertyException Constructor.
        /// </summary>
        /// <param name="message">Error message text.</param>
        /// <param name="name">The name of the property that is at the origin of the exception.</param>
        /// <param name="innerException">Inner exception.</param>
        public PropertyException(
            string message,
            string name,
            Exception innerException)
            : base(message, innerException)
        {
            this.name = name;
        }

        /// <summary>
        /// Gets the name of the property that caused the exception.
        /// </summary>
        public string Name
        {
            get { return this.name; }
        }
    }
}
