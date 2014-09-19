// ---------------------------------------------------------------------------
// <copyright file="ISearchStringProvider.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the ISearchStringProvider interface.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    /// <summary>
    /// Interface defined for types that can produce a string representation for use in search filters.
    /// </summary>
    public interface ISearchStringProvider
    {
        /// <summary>
        /// Get a string representation for using this instance in a search filter.
        /// </summary>
        /// <returns>String representation of instance.</returns>
        string GetSearchString();
    }
}
