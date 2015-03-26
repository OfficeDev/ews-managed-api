// ---------------------------------------------------------------------------
// <copyright file="FindPeopleResults.cs" company="Microsoft">
//     Copyright (c) Microsoft Corporation.  All rights reserved.
// </copyright>
// ---------------------------------------------------------------------------

//-----------------------------------------------------------------------
// <summary>Defines the FindPeopleResults class.</summary>
//-----------------------------------------------------------------------

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the results of an Persona search operation.
    /// </summary>
    public sealed class FindPeopleResults
    {
        /// <summary>
        /// Creates a new instance of the <see cref="FindPeopleResults"/> class.
        /// </summary>
        internal FindPeopleResults()
        {
            this.Personas = new Collection<Persona>();
        }

        /// <summary>
        /// Accessors for the Personas that were found by the search operation.
        /// </summary>
        public Collection<Persona> Personas { get; internal set; }

        /// <summary>
        /// Accessors for the total count of Personas in view.
        /// </summary>
        public int? TotalCount { get; internal set; }

        /// <summary>
        /// Accessors for the first index of the matching row.
        /// </summary>
        public int? FirstMatchingRowIndex { get; internal set; }

        /// <summary>
        /// Accessors for the first index of the loaded row.
        /// </summary>
        public int? FirstLoadedRowIndex { get; internal set; }
    }
}
